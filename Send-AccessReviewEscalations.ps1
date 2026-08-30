<#
.SYNOPSIS
    Monitors active M365 Group access reviews and emails fallback admins
    about reviews nearing their deadline with unreviewed members.

.DESCRIPTION
    Connects to Microsoft Graph (interactive login), finds all access review
    definitions created by the New-M365GroupAccessReviews.ps1 script (named
    "Access Review - <GroupName>"), and for each currently active instance:
      - Calculates days remaining until the review's end date
      - Checks whether any members are still "NotReviewed"
      - If within the escalation window AND unreviewed members remain,
        emails the review's fallback admin group (resolved from the
        review definition itself - no CSV needed) with a summary.

    Intended to run on a schedule (e.g. daily, via Task Scheduler) a few
    days before each 14-day review window closes.

.NOTES
    Requires modules: Microsoft.Graph.Authentication,
                       Microsoft.Graph.Identity.Governance,
                       Microsoft.Graph.Users.Actions
    Required Graph scopes: AccessReview.Read.All, Group.Read.All, Mail.Send
    Mail.Send sends from the signed-in account's mailbox (delegated).

.PARAMETER DaysBeforeDeadline
    How many days before a review instance's end date to start escalating.
    Default: 3

.PARAMETER DryRun
    If specified, shows what WOULD be escalated/emailed without actually
    sending any mail. Recommended for the first run.

.EXAMPLE
    .\Send-AccessReviewEscalations.ps1 -DryRun
    .\Send-AccessReviewEscalations.ps1 -DaysBeforeDeadline 2
#>

[CmdletBinding()]
param(
    [int]$DaysBeforeDeadline = 3,
    [switch]$DryRun
)

# ---------------------------------------------------------------------------
# 1. Ensure required modules are present
# ---------------------------------------------------------------------------
$requiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Identity.Governance',
    'Microsoft.Graph.Users.Actions'
)

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing module: $module ..." -ForegroundColor Yellow
        Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module -Name $module -ErrorAction Stop
}

# ---------------------------------------------------------------------------
# 2. Connect to Microsoft Graph (interactive login)
# ---------------------------------------------------------------------------
Write-Host "Connecting to Microsoft Graph (interactive login)..." -ForegroundColor Cyan

Connect-MgGraph -Scopes "AccessReview.Read.All", "Group.Read.All", "Mail.Send" -NoWelcome

$context = Get-MgContext
Write-Host "Connected as: $($context.Account) | Tenant: $($context.TenantId)" -ForegroundColor Green

if ($DryRun) {
    Write-Host "*** DRY RUN MODE - no emails will be sent ***" -ForegroundColor Magenta
}

# ---------------------------------------------------------------------------
# 3. Helper: cached resolution of group ID -> group object (for names/emails)
# ---------------------------------------------------------------------------
$groupCache = @{}

function Get-CachedGroup {
    param([string]$GroupId)

    if ($groupCache.ContainsKey($GroupId)) {
        return $groupCache[$GroupId]
    }

    try {
        $grp = Get-MgGroup -GroupId $GroupId -Property Id, DisplayName, Mail -ErrorAction Stop
        $groupCache[$GroupId] = $grp
        return $grp
    }
    catch {
        Write-Warning "Could not resolve group '$GroupId': $($_.Exception.Message)"
        $groupCache[$GroupId] = $null
        return $null
    }
}

# ---------------------------------------------------------------------------
# 4. Retrieve all access review definitions created by our creation script
# ---------------------------------------------------------------------------
Write-Host "Retrieving access review definitions..." -ForegroundColor Cyan

$definitions = Get-MgIdentityGovernanceAccessReviewDefinition -All `
    -Filter "startswith(displayName,'Access Review - ')"

Write-Host "Found $($definitions.Count) access review definition(s)." -ForegroundColor Green

# ---------------------------------------------------------------------------
# 5. Walk each definition's active instances and evaluate for escalation
# ---------------------------------------------------------------------------
$escalations = New-Object System.Collections.Generic.List[Object]
$counter = 0

foreach ($def in $definitions) {
    $counter++
    Write-Progress -Activity "Checking access reviews" `
        -Status "$counter of $($definitions.Count): $($def.DisplayName)" `
        -PercentComplete (($counter / $definitions.Count) * 100)

    # Dump to JSON once and regex out the group IDs - more reliable than
    # relying on the exact strong-typed model shape returned by the SDK.
    $defJson = $def | ConvertTo-Json -Depth 20

    $scopeGroupId = $null
    if ($defJson -match '/groups/([0-9a-fA-F-]{36})/transitiveMembers') {
        $scopeGroupId = $Matches[1]
    }

    $fallbackGroupId = $null
    if ($defJson -match '/groups/([0-9a-fA-F-]{36})/members') {
        $fallbackGroupId = $Matches[1]
    }

    $reviewedGroup = if ($scopeGroupId) { Get-CachedGroup -GroupId $scopeGroupId } else { $null }
    $adminGroup    = if ($fallbackGroupId) { Get-CachedGroup -GroupId $fallbackGroupId } else { $null }

    if (-not $adminGroup -or -not $adminGroup.Mail) {
        Write-Warning "Definition '$($def.DisplayName)': could not resolve a fallback admin group email - skipping escalation check."
        continue
    }

    # --- Get active instances for this definition -----------------------
    try {
        $instances = Get-MgIdentityGovernanceAccessReviewDefinitionInstance `
            -AccessReviewScheduleDefinitionId $def.Id -All -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not retrieve instances for '$($def.DisplayName)': $($_.Exception.Message)"
        continue
    }

    $activeInstances = $instances | Where-Object { $_.Status -eq 'InProgress' }

    foreach ($instance in $activeInstances) {
        if (-not $instance.EndDateTime) { continue }

        $daysRemaining = [math]::Ceiling(($instance.EndDateTime - (Get-Date).ToUniversalTime()).TotalDays)

        if ($daysRemaining -gt $DaysBeforeDeadline -or $daysRemaining -lt 0) {
            continue  # Not yet in the escalation window
        }

        # --- Check for unreviewed decisions ------------------------------
        try {
            $decisions = Get-MgIdentityGovernanceAccessReviewDefinitionInstanceDecision `
                -AccessReviewScheduleDefinitionId $def.Id `
                -AccessReviewInstanceId $instance.Id `
                -All -ErrorAction Stop
        }
        catch {
            Write-Warning "Could not retrieve decisions for '$($def.DisplayName)' instance '$($instance.Id)': $($_.Exception.Message)"
            continue
        }

        $notReviewedCount = ($decisions | Where-Object { $_.ReviewResult -eq 'NotReviewed' }).Count
        $totalCount        = $decisions.Count

        if ($notReviewedCount -eq 0) {
            continue  # Fully reviewed already, nothing to escalate
        }

        $groupName = if ($reviewedGroup) { $reviewedGroup.DisplayName } else { $def.DisplayName -replace '^Access Review - ', '' }

        $escalations.Add([PSCustomObject]@{
            GroupDisplayName  = $groupName
            DefinitionId      = $def.Id
            InstanceId        = $instance.Id
            EndDateTime       = $instance.EndDateTime
            DaysRemaining     = $daysRemaining
            NotReviewedCount  = $notReviewedCount
            TotalMembers      = $totalCount
            AdminGroupEmail   = $adminGroup.Mail
            AdminGroupName    = $adminGroup.DisplayName
        })
    }
}

Write-Progress -Activity "Checking access reviews" -Completed

Write-Host "`nFound $($escalations.Count) review(s) needing escalation." -ForegroundColor $(if ($escalations.Count -gt 0) { 'Yellow' } else { 'Green' })

# ---------------------------------------------------------------------------
# 6. Send escalation emails (grouped by admin group, one email per review)
# ---------------------------------------------------------------------------
$sentCount = 0

foreach ($esc in $escalations) {
    $subject = "Action needed: Access Review for '$($esc.GroupDisplayName)' closes in $($esc.DaysRemaining) day(s)"

    $body = @"
Hello,

The access review for the group <b>$($esc.GroupDisplayName)</b> is closing soon and still has unreviewed members.

<ul>
  <li><b>Review closes:</b> $($esc.EndDateTime.ToString('yyyy-MM-dd HH:mm')) UTC ($($esc.DaysRemaining) day(s) remaining)</li>
  <li><b>Unreviewed members:</b> $($esc.NotReviewedCount) of $($esc.TotalMembers)</li>
</ul>

As the fallback reviewer group for this access review, please log into the Entra admin center (Identity Governance &gt; Access Reviews) and complete the outstanding decisions before the review closes.

This is an automated notification.
"@

    if ($DryRun) {
        Write-Host "[DRY RUN] Would email $($esc.AdminGroupEmail) about '$($esc.GroupDisplayName)' ($($esc.NotReviewedCount) unreviewed, $($esc.DaysRemaining) day(s) left)" -ForegroundColor Magenta
        continue
    }

    try {
        $message = @{
            subject      = $subject
            body         = @{
                contentType = "HTML"
                content     = $body
            }
            toRecipients = @(
                @{ emailAddress = @{ address = $esc.AdminGroupEmail } }
            )
        }

        Send-MgUserMail -UserId "me" -Message $message -SaveToSentItems -ErrorAction Stop
        Write-Host "Emailed $($esc.AdminGroupEmail) about '$($esc.GroupDisplayName)'." -ForegroundColor Green
        $sentCount++
    }
    catch {
        Write-Warning "Failed to email $($esc.AdminGroupEmail) for '$($esc.GroupDisplayName)': $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# 7. Summary + log
# ---------------------------------------------------------------------------
Write-Host "`n----- Summary -----" -ForegroundColor Cyan
Write-Host "Reviews needing escalation: $($escalations.Count)" -ForegroundColor Yellow
if (-not $DryRun) {
    Write-Host "Emails sent: $sentCount" -ForegroundColor Green
}

if ($escalations.Count -gt 0) {
    $logPath = ".\AccessReview_Escalations_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $escalations | Export-Csv -Path $logPath -NoTypeInformation
    Write-Host "Escalation log saved to: $((Resolve-Path $logPath).Path)" -ForegroundColor Cyan
}

# ---------------------------------------------------------------------------
# 8. Disconnect
# ---------------------------------------------------------------------------
Disconnect-MgGraph | Out-Null
