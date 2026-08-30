<#
.SYNOPSIS
    Bulk-creates Microsoft Entra ID Access Reviews for M365 Groups from a CSV.

.DESCRIPTION
    Reads a CSV (GroupId, GroupDisplayName, Admins) and creates one Access
    Review Schedule Definition per group:
      - Reviewer: dynamic query -> the group's current owner(s) at review time
      - Fallback reviewer: only triggers when a group has NO owners; uses
        the individual admin UPNs listed in the "Admins" column
      - Recurrence: semi-annually, 14-day review window, no end date
      - No auto-apply of decisions (admin reviews and actions manually)
      - No response after 14 days = auto-approved
      - Self-review disabled

.NOTES
    Requires module: Microsoft.Graph.Identity.Governance
    Required Graph scopes: AccessReview.ReadWrite.All, Group.Read.All,
                            User.Read.All
    License: Microsoft Entra ID Governance (P2) required in the tenant.

.PARAMETER CsvPath
    Path to the input CSV with columns: GroupId, GroupDisplayName, Admins
    The Admins column holds the admin (mail-enabled) group's EMAIL ADDRESS.
    Its current members become the fallback reviewers dynamically - only
    used when a group has zero owners. Resolved to Object ID at runtime.

.PARAMETER AdminDelimiter
    Delimiter used if multiple admin group emails are listed in the Admins
    column. Defaults to semicolon (;).

.EXAMPLE
    .\New-M365GroupAccessReviews.ps1 -CsvPath ".\M365Groups.csv"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CsvPath,

    [string]$AdminDelimiter = ";"
)

# ---------------------------------------------------------------------------
# 1. Ensure required module is present
# ---------------------------------------------------------------------------
$requiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Identity.Governance'
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

Connect-MgGraph -Scopes "AccessReview.ReadWrite.All", "Group.Read.All", "User.Read.All" -NoWelcome

$context = Get-MgContext
Write-Host "Connected as: $($context.Account) | Tenant: $($context.TenantId)" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 3. Load CSV
# ---------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) {
    if (Test-Path -LiteralPath $CsvPath -PathType Container) {
        throw "The path '$CsvPath' is a folder, not a file. Point -CsvPath at the actual .csv file inside it, e.g. '$CsvPath\M365Groups.csv'."
    }
    throw "CSV file not found: $CsvPath"
}

$groups = Import-Csv -Path $CsvPath

# Validate required columns
$requiredColumns = @('GroupId', 'GroupDisplayName', 'Admins')
foreach ($col in $requiredColumns) {
    if (-not ($groups | Get-Member -Name $col -MemberType NoteProperty)) {
        throw "CSV is missing required column: $col"
    }
}

Write-Host "Loaded $($groups.Count) group(s) from CSV." -ForegroundColor Green

# ---------------------------------------------------------------------------
# 4. Helper: cached resolution of admin group email -> Object ID
# ---------------------------------------------------------------------------
$adminGroupIdCache = @{}

function Resolve-AdminGroupId {
    param([string]$Email)

    if ($adminGroupIdCache.ContainsKey($Email)) {
        return $adminGroupIdCache[$Email]
    }

    try {
        $adminGroup = Get-MgGroup -Filter "mail eq '$Email'" -ErrorAction Stop
        if (-not $adminGroup) {
            Write-Warning "No group found with email '$Email'."
            $adminGroupIdCache[$Email] = $null
            return $null
        }
        if ($adminGroup.Count -gt 1) {
            Write-Warning "Multiple groups found with email '$Email' - using the first match."
            $adminGroup = $adminGroup[0]
        }
        $adminGroupIdCache[$Email] = $adminGroup.Id
        return $adminGroup.Id
    }
    catch {
        Write-Warning "Failed to resolve admin group '$Email': $($_.Exception.Message)"
        $adminGroupIdCache[$Email] = $null
        return $null
    }
}

# ---------------------------------------------------------------------------
# 5. Build the recurring schedule / settings (shared across all reviews)
# ---------------------------------------------------------------------------
$startDate = (Get-Date).ToString("yyyy-MM-dd")

$commonSettings = @{
    mailNotificationsEnabled          = $true
    reminderNotificationsEnabled      = $true
    justificationRequiredOnApproval   = $true
    defaultDecisionEnabled            = $true
    defaultDecision                   = "Approve"      # No-response = auto-approved
    instanceDurationInDays            = 14              # Review window length
    autoApplyDecisionsEnabled         = $false           # Record decisions only; admin acts manually
    recommendationsEnabled            = $true
    recurrence                        = @{
        pattern = @{
            type     = "absoluteMonthly"
            interval = 6                                # Semi-annually
        }
        range = @{
            type      = "noEnd"
            startDate = $startDate
        }
    }
}

# ---------------------------------------------------------------------------
# 6. Loop through groups and create an access review for each
# ---------------------------------------------------------------------------
$createdReviews = New-Object System.Collections.Generic.List[Object]
$skippedReviews = New-Object System.Collections.Generic.List[Object]
$failedReviews  = New-Object System.Collections.Generic.List[Object]
$counter = 0

foreach ($row in $groups) {
    $counter++
    $groupId   = $row.GroupId.Trim()
    $groupName = $row.GroupDisplayName.Trim()
    $reviewDisplayName = "Access Review - $groupName"

    Write-Progress -Activity "Creating Access Reviews" `
        -Status "$counter of $($groups.Count): $groupName" `
        -PercentComplete (($counter / $groups.Count) * 100)

    if ([string]::IsNullOrWhiteSpace($groupId)) {
        Write-Warning "Row $counter has no GroupId, skipping: $groupName"
        $failedReviews.Add([PSCustomObject]@{ GroupDisplayName = $groupName; GroupId = $groupId; Reason = "Missing GroupId" })
        continue
    }

    # --- Duplicate protection: skip if a review with the same name exists --
    try {
        $existing = Get-MgIdentityGovernanceAccessReviewDefinition `
            -Filter "displayName eq '$reviewDisplayName'" -ErrorAction Stop
    }
    catch {
        $existing = $null
    }

    if ($existing) {
        Write-Host "Skipping '$groupName' - access review already exists." -ForegroundColor Yellow
        $skippedReviews.Add([PSCustomObject]@{ GroupDisplayName = $groupName; GroupId = $groupId; Reason = "Already exists" })
        continue
    }

    # --- Parse admin group email(s) from the Admins column, resolve to ID --
    $adminEmails = @()
    if (-not [string]::IsNullOrWhiteSpace($row.Admins)) {
        $adminEmails = $row.Admins -split [regex]::Escape($AdminDelimiter) |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -ne "" }
    }

    if ($adminEmails.Count -eq 0) {
        Write-Warning "Row $counter ('$groupName') has no admin group email listed - fallback reviewers will be empty."
    }

    $fallbackReviewers = @()
    foreach ($email in $adminEmails) {
        $adminGroupId = Resolve-AdminGroupId -Email $email
        if ($adminGroupId) {
            # Current members of the admin group become fallback reviewers dynamically
            $fallbackReviewers += @{
                query     = "/groups/$adminGroupId/members"
                queryType = "MicrosoftGraph"
            }
        }
        else {
            Write-Warning "Row $counter ('$groupName'): could not resolve admin group '$email' - skipped from fallback reviewers."
        }
    }

    # --- Build the request body ---------------------------------------------
    $body = @{
        displayName            = $reviewDisplayName
        descriptionForAdmins   = "Recurring access review of membership for group '$groupName'. Reviewed by group owner(s); falls back to designated admins if no owner exists."
        descriptionForReviewers = "Please review who should retain access to the group '$groupName'."
        scope                   = @{
            "@odata.type" = "#microsoft.graph.accessReviewQueryScope"
            query         = "/groups/$groupId/transitiveMembers"
            queryType     = "MicrosoftGraph"
        }
        reviewers = @(
            @{
                query     = "/groups/$groupId/owners"
                queryType = "MicrosoftGraph"
            }
        )
        fallbackReviewers = $fallbackReviewers
        settings           = $commonSettings
    }

    # --- Create the review ---------------------------------------------------
    try {
        $newReview = New-MgIdentityGovernanceAccessReviewDefinition -BodyParameter $body -ErrorAction Stop
        Write-Host "Created access review for '$groupName'." -ForegroundColor Green
        $createdReviews.Add([PSCustomObject]@{
            GroupDisplayName = $groupName
            GroupId          = $groupId
            ReviewId         = $newReview.Id
            AdminFallbackGroups = ($adminEmails -join "; ")
        })
    }
    catch {
        Write-Warning "Failed to create access review for '$groupName': $($_.Exception.Message)"
        $failedReviews.Add([PSCustomObject]@{ GroupDisplayName = $groupName; GroupId = $groupId; Reason = $_.Exception.Message })
    }
}

Write-Progress -Activity "Creating Access Reviews" -Completed

# ---------------------------------------------------------------------------
# 6. Summary
# ---------------------------------------------------------------------------
Write-Host "`n----- Summary -----" -ForegroundColor Cyan
Write-Host "Created: $($createdReviews.Count)" -ForegroundColor Green
Write-Host "Skipped (already existed): $($skippedReviews.Count)" -ForegroundColor Yellow
Write-Host "Failed: $($failedReviews.Count)" -ForegroundColor Red

if ($failedReviews.Count -gt 0) {
    Write-Host "`nFailed rows:" -ForegroundColor Red
    $failedReviews | Format-Table -AutoSize
}

# Export a run log for reference
$logPath = ".\AccessReview_CreationLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$allResults = @()
$allResults += $createdReviews | Select-Object *, @{N='Status';E={'Created'}}
$allResults += $skippedReviews | Select-Object *, @{N='Status';E={'Skipped'}}
$allResults += $failedReviews  | Select-Object *, @{N='Status';E={'Failed'}}
$allResults | Export-Csv -Path $logPath -NoTypeInformation

Write-Host "`nRun log saved to: $((Resolve-Path $logPath).Path)" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# 7. Disconnect
# ---------------------------------------------------------------------------
Disconnect-MgGraph | Out-Null
