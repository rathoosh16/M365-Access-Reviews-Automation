<#
.SYNOPSIS
    Exports all Microsoft 365 Groups (with owners, member count, Team,
    SharePoint site status, and owner account health) to an Excel file.

.DESCRIPTION
    Connects to Microsoft Graph AND Exchange Online (both interactive login),
    retrieves all M365 (Unified) Groups, and writes ONE ROW PER GROUP to
    Excel. Owners are combined into a single semicolon-separated column.
    For every owner, checks:
      - whether the account is Disabled (Entra ID AccountEnabled = false)
      - whether the account is a Shared Mailbox (Exchange RecipientTypeDetails)
    Groups with a problem owner (disabled or shared mailbox) are flagged so
    they're easy to filter/prioritise for access reviews.

.NOTES
    Requires modules: Microsoft.Graph, ExchangeOnlineManagement, ImportExcel
    Required Graph scopes: Group.Read.All, User.Read.All, Directory.Read.All,
                            Sites.Read.All
    Exchange Online: read-only (View-Only Recipients) is enough.

.EXAMPLE
    .\Get-M365GroupsAndOwners.ps1
    .\Get-M365GroupsAndOwners.ps1 -OutputPath "C:\Reports\M365Groups.xlsx"
#>

[CmdletBinding()]
param(
    [string]$OutputPath = ".\M365Groups_Owners_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
)

# ---------------------------------------------------------------------------
# 1. Ensure required modules are present
# ---------------------------------------------------------------------------
$requiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Sites',
    'ExchangeOnlineManagement',
    'ImportExcel'
)

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing module: $module ..." -ForegroundColor Yellow
        Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module -Name $module -ErrorAction Stop
}

# ---------------------------------------------------------------------------
# 2. Connect to Microsoft Graph and Exchange Online (interactive login)
# ---------------------------------------------------------------------------
Write-Host "Connecting to Microsoft Graph (interactive login)..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "Group.Read.All", "User.Read.All", "Directory.Read.All", "Sites.Read.All" -NoWelcome

$context = Get-MgContext
Write-Host "Connected to Graph as: $($context.Account) | Tenant: $($context.TenantId)" -ForegroundColor Green

Write-Host "Connecting to Exchange Online (interactive login)..." -ForegroundColor Cyan
Connect-ExchangeOnline -ShowBanner:$false

# ---------------------------------------------------------------------------
# 3. Retrieve all Microsoft 365 (Unified) Groups
# ---------------------------------------------------------------------------
Write-Host "Retrieving Microsoft 365 groups..." -ForegroundColor Cyan

$m365Groups = Get-MgGroup -All -Filter "groupTypes/any(c:c eq 'Unified')" `
    -Property Id, DisplayName, Mail, Description, Visibility, CreatedDateTime, MembershipRule, GroupTypes, ResourceProvisioningOptions

# Safety: de-duplicate on Group Id in case Graph paging ever returns a repeat
$m365Groups = $m365Groups | Sort-Object Id -Unique

Write-Host "Found $($m365Groups.Count) unique Microsoft 365 groups." -ForegroundColor Green

# ---------------------------------------------------------------------------
# 4. Helper: cached lookup of owner account status (Entra + Exchange)
# ---------------------------------------------------------------------------
$ownerStatusCache = @{}

function Get-OwnerStatus {
    param(
        [string]$OwnerId,
        [string]$OwnerUPN
    )

    if ($ownerStatusCache.ContainsKey($OwnerId)) {
        return $ownerStatusCache[$OwnerId]
    }

    $accountEnabled = $null
    $isSharedMailbox = $false
    $recipientTypeDetails = $null

    # --- Entra ID account status ---
    try {
        $userObj = Get-MgUser -UserId $OwnerId -Property Id, DisplayName, UserPrincipalName, AccountEnabled -ErrorAction Stop
        $accountEnabled = $userObj.AccountEnabled
    }
    catch {
        # Owner could be a non-user object (e.g. a deleted account, or a group) - treat as unknown
        $accountEnabled = "Unknown"
    }

    # --- Exchange mailbox type (Shared Mailbox check) ---
    try {
        $mbx = Get-EXOMailbox -Identity $OwnerUPN -Properties RecipientTypeDetails -ErrorAction Stop
        $recipientTypeDetails = $mbx.RecipientTypeDetails
        if ($recipientTypeDetails -eq 'SharedMailbox') {
            $isSharedMailbox = $true
        }
    }
    catch {
        # No mailbox found / not accessible - leave as null
        $recipientTypeDetails = "N/A"
    }

    $status = [PSCustomObject]@{
        AccountEnabled        = $accountEnabled
        IsSharedMailbox       = $isSharedMailbox
        RecipientTypeDetails  = $recipientTypeDetails
    }

    $ownerStatusCache[$OwnerId] = $status
    return $status
}

# ---------------------------------------------------------------------------
# 5. Loop through each group once, gather owners (+ status), members, Team, SPO
# ---------------------------------------------------------------------------
$results = New-Object System.Collections.Generic.List[Object]
$counter = 0

foreach ($group in $m365Groups) {
    $counter++
    Write-Progress -Activity "Processing M365 Groups" `
        -Status "$counter of $($m365Groups.Count): $($group.DisplayName)" `
        -PercentComplete (($counter / $m365Groups.Count) * 100)

    # --- Owners + status -------------------------------------------------
    try {
        $owners = Get-MgGroupOwner -GroupId $group.Id -All -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not retrieve owners for group '$($group.DisplayName)': $($_.Exception.Message)"
        $owners = @()
    }

    $ownerNameList   = @()
    $ownerUPNList    = @()
    $ownerFlagList   = @()   # per-owner human-readable flag, e.g. "Jane Doe (Disabled)"
    $anyOwnerDisabled       = $false
    $anyOwnerSharedMailbox  = $false

    foreach ($owner in $owners) {
        $ownerId          = $owner.Id
        $ownerDisplayName = $owner.AdditionalProperties['displayName']
        $ownerUPN         = $owner.AdditionalProperties['userPrincipalName']

        $ownerNameList += $ownerDisplayName
        $ownerUPNList  += $ownerUPN

        if ($ownerUPN) {
            $status = Get-OwnerStatus -OwnerId $ownerId -OwnerUPN $ownerUPN

            $issueTags = @()
            if ($status.AccountEnabled -eq $false) {
                $issueTags += "Disabled"
                $anyOwnerDisabled = $true
            }
            if ($status.IsSharedMailbox) {
                $issueTags += "Shared Mailbox"
                $anyOwnerSharedMailbox = $true
            }

            if ($issueTags.Count -gt 0) {
                $ownerFlagList += "$ownerDisplayName ($($issueTags -join ', '))"
            }
        }
    }

    if ($owners.Count -eq 0) {
        $ownerNamesJoined = "*** NO OWNER ***"
        $ownerUPNsJoined  = ""
    }
    else {
        $ownerNamesJoined = $ownerNameList -join "; "
        $ownerUPNsJoined  = $ownerUPNList -join "; "
    }

    $ownerIssuesJoined = if ($ownerFlagList.Count -gt 0) { $ownerFlagList -join "; " } else { "" }

    # --- Member count -----------------------------------------------------
    try {
        $memberCount = (Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop).Count
    }
    catch {
        Write-Warning "Could not retrieve members for group '$($group.DisplayName)': $($_.Exception.Message)"
        $memberCount = "Error"
    }

    # --- Team provisioned? --------------------------------------------------
    $hasTeam = if ($group.ResourceProvisioningOptions -contains "Team") { "Yes" } else { "No" }

    # --- SharePoint site exists? ---------------------------------------------
    try {
        $null = Get-MgGroupSite -GroupId $group.Id -SiteId "root" -ErrorAction Stop
        $hasSharePointSite = "Yes"
    }
    catch {
        $hasSharePointSite = "No"
    }

    $results.Add([PSCustomObject]@{
        GroupDisplayName      = $group.DisplayName
        GroupEmail            = $group.Mail
        GroupId               = $group.Id
        Visibility            = $group.Visibility
        MemberCount           = $memberCount
        HasTeam               = $hasTeam
        HasSharePointSite     = $hasSharePointSite
        Owners                = $ownerNamesJoined
        OwnerUPNs             = $ownerUPNsJoined
        AnyOwnerDisabled      = if ($owners.Count -eq 0) { "N/A" } elseif ($anyOwnerDisabled) { "Yes" } else { "No" }
        AnyOwnerSharedMailbox = if ($owners.Count -eq 0) { "N/A" } elseif ($anyOwnerSharedMailbox) { "Yes" } else { "No" }
        OwnerIssueDetails     = $ownerIssuesJoined
        MembershipRule        = $group.MembershipRule
        Description           = $group.Description
        CreatedDateTime       = $group.CreatedDateTime
    })
}

Write-Progress -Activity "Processing M365 Groups" -Completed

# ---------------------------------------------------------------------------
# 6. Export to Excel
# ---------------------------------------------------------------------------
Write-Host "Exporting $($results.Count) rows to Excel: $OutputPath" -ForegroundColor Cyan

$results |
    Sort-Object GroupDisplayName |
    Export-Excel -Path $OutputPath `
        -WorksheetName "M365 Groups" `
        -AutoSize `
        -AutoFilter `
        -FreezeTopRow `
        -BoldTopRow `
        -TableName "M365Groups"

# Summary sheet 1: groups with no owners at all
$noOwnerGroups = $results | Where-Object { $_.Owners -eq "*** NO OWNER ***" }
if ($noOwnerGroups.Count -gt 0) {
    $noOwnerGroups |
        Select-Object GroupDisplayName, GroupEmail, GroupId, Visibility, MemberCount, CreatedDateTime |
        Export-Excel -Path $OutputPath `
            -WorksheetName "Groups Without Owners" `
            -AutoSize -AutoFilter -BoldTopRow
}

# Summary sheet 2: groups where at least one owner is disabled or a shared mailbox
$problemOwnerGroups = $results | Where-Object { $_.AnyOwnerDisabled -eq "Yes" -or $_.AnyOwnerSharedMailbox -eq "Yes" }
if ($problemOwnerGroups.Count -gt 0) {
    $problemOwnerGroups |
        Select-Object GroupDisplayName, GroupEmail, GroupId, Owners, OwnerIssueDetails, AnyOwnerDisabled, AnyOwnerSharedMailbox |
        Export-Excel -Path $OutputPath `
            -WorksheetName "Owners - Disabled or Shared" `
            -AutoSize -AutoFilter -BoldTopRow
}

Write-Host "Done. File saved to: $((Resolve-Path $OutputPath).Path)" -ForegroundColor Green
Write-Host "Total groups: $($results.Count) | No owner: $($noOwnerGroups.Count) | Owner issue (disabled/shared): $($problemOwnerGroups.Count)" -ForegroundColor Yellow

# ---------------------------------------------------------------------------
# 7. Disconnect
# ---------------------------------------------------------------------------
Disconnect-MgGraph | Out-Null
Disconnect-ExchangeOnline -Confirm:$false
