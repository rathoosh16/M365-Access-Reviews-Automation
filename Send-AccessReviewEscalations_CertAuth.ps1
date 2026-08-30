<#
.SYNOPSIS
    Monitors active M365 Group access reviews and emails fallback admins
    about reviews nearing their deadline with unreviewed members.

.DESCRIPTION
    Connects to Microsoft Graph using Certificate Authentication (Service Principal),
    finds access review definitions ("Access Review - <GroupName>"), and evaluates
    in-progress reviews requiring escalation before deadlines. Designed for bi-weekly 
    execution via Task Scheduler.

.NOTES
    Required Microsoft Graph Application Permissions:
      - AccessReview.Read.All
      - Group.Read.All
      - Mail.Send (Granted to Application)

.PARAMETER DaysBeforeDeadline
    How many days before a review instance's end date to start escalating. 
    Default: 14 (Matches a 2-week scheduled task frequency).

.PARAMETER DryRun
    If specified, logs what WOULD be escalated/emailed without sending mail.
#>

[CmdletBinding()]
param(
    [int]$DaysBeforeDeadline = 14,
    [switch]$DryRun
)

# ---------------------------------------------------------------------------
# 0. HARDCODED TENANT & AUTHENTICATION PARAMETERS
# ---------------------------------------------------------------------------
$TenantId              = "company.com"               # Replace with your Directory GUID if domain resolution isn't used
$ClientId              = "YOUR_APP_REGISTRATION_CLIENT_ID" # Place your Entra App Registration Client ID here
$CertificateThumbprint = "YOUR_CERTIFICATE_THUMBPRINT"    # Place your installed Certificate Thumbprint here
$SenderMailbox         = "sender@company.comh"        # Sending mailbox address

# ---------------------------------------------------------------------------
# 1. Setup Logging & Output Directories
# ---------------------------------------------------------------------------
$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$LogFolder = Join-Path -Path $ScriptRoot -ChildPath "Logs"
$ReportFolder = Join-Path -Path $ScriptRoot -ChildPath "Reports"

ForEach ($Folder in @($LogFolder, $ReportFolder)) {
    If (-not (Test-Path $Folder)) { New-Item -Path $Folder -ItemType Directory -Force | Out-Null }
}

$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$TranscriptPath = Join-Path -Path $LogFolder -ChildPath "AccessReview_Escalation_$TimeStamp.log"
Start-Transcript -Path $TranscriptPath -NoClobber | Out-Null

# ---------------------------------------------------------------------------
# 2. Ensure required modules are present
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
# 3. Connect to Microsoft Graph (Hardcoded Certificate Auth)
# ---------------------------------------------------------------------------
Write-Host "Connecting to Microsoft Graph via Certificate Authentication..." -ForegroundColor Cyan

Try {
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorAction Stop
} Catch {
    Write-Host "ERROR: Connection failed - $($_.Exception.Message)" -ForegroundColor Red
    Stop-Transcript | Out-Null
    exit
}

$context = Get-MgContext
Write-Host "Connected successfully as App ID: $($context.ClientId) | Tenant: $($context.TenantId)" -ForegroundColor Green

if ($DryRun) {
    Write-Host "*** DRY RUN MODE ENABLED - No emails will be sent ***" -ForegroundColor Magenta
}

# ---------------------------------------------------------------------------
# 4. Helper: cached resolution of group ID -> group object
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
# 5. Retrieve access review definitions
# ---------------------------------------------------------------------------
Write-Host "`nRetrieving access review definitions..." -ForegroundColor Cyan

$definitions = Get-MgIdentityGovernanceAccessReviewDefinition -All `
    -Filter "startswith(displayName,'Access Review - ')"

Write-Host "Found $($definitions.Count) access review definition(s)." -ForegroundColor Green

# ---------------------------------------------------------------------------
# 6. Walk each definition's active instances and evaluate for escalation
# ---------------------------------------------------------------------------
$escalations = New-Object System.Collections.Generic.List[Object]
$counter = 0

foreach ($def in $definitions) {
    $counter++
    Write-Progress -Activity "Checking access reviews" `
        -Status "$counter of $($definitions.Count): $($def.DisplayName)" `
        -PercentComplete (($counter / $definitions.Count) * 100)

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
            continue  # Not inside the target escalation window
        }

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
            continue  # Fully reviewed
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
# 7. Send escalation emails
# ---------------------------------------------------------------------------
$sentCount = 0

foreach ($esc in $escalations) {
    $subject = "Action Needed: Access Review for '$($esc.GroupDisplayName)' closes in $($esc.DaysRemaining) day(s)"

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

        # Send-MgUserMail using defined service mailbox
        Send-MgUserMail -UserId $SenderMailbox -Message $message -SaveToSentItems -ErrorAction Stop
        Write-Host "Emailed $($esc.AdminGroupEmail) about '$($esc.GroupDisplayName)'." -ForegroundColor Green
        $sentCount++
    }
    catch {
        Write-Warning "Failed to email $($esc.AdminGroupEmail) for '$($esc.GroupDisplayName)': $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# 8. Summary + Log Export
# ---------------------------------------------------------------------------
Write-Host "`n----- Summary -----" -ForegroundColor Cyan
Write-Host "Reviews needing escalation: $($escalations.Count)" -ForegroundColor Yellow
if (-not $DryRun) {
    Write-Host "Emails sent: $sentCount" -ForegroundColor Green
}

if ($escalations.Count -gt 0) {
    $ReportCSVPath = Join-Path -Path $ReportFolder -ChildPath "AccessReview_Escalations_$TimeStamp.csv"
    $escalations | Export-Csv -Path $ReportCSVPath -NoTypeInformation
    Write-Host "Escalation log saved to: $ReportCSVPath" -ForegroundColor Cyan
}

# ---------------------------------------------------------------------------
# 9. Cleanup & Disconnect
# ---------------------------------------------------------------------------
Disconnect-MgGraph | Out-Null
Stop-Transcript | Out-Null