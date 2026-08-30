# M365 Group Access Reviews Automation

Automate the Microsoft 365 Group access-review lifecycle with PowerShell and Microsoft Graph — from group discovery and owner-health analysis to recurring Access Review creation and scheduled escalation.

> **Current scope:** This project currently covers **Microsoft 365 Groups only**. Additional workloads such as Microsoft Teams, SharePoint, Applications, Service Principals, and other Entra resources may be added in future iterations.

## What this project does

The solution is designed around three stages:

1. **Discover and assess M365 Groups**
   - Enumerates Microsoft 365 (Unified) Groups.
   - Collects group owners, member count, Team provisioning status, and SharePoint site status.
   - Checks owner account health in Entra ID.
   - Checks whether an owner is a Shared Mailbox in Exchange Online.
   - Highlights groups with no owners or problematic owners.
   - Exports the results to Excel.

2. **Create recurring Entra ID Access Reviews**
   - Creates one Access Review definition per M365 Group from a CSV.
   - Reviews the group's transitive membership.
   - Uses the group's current owners as reviewers.
   - Uses designated admin-group members as fallback reviewers when the target group has no owners.
   - Recurs every 6 months.
   - Each review instance remains open for 14 days.
   - Recommendations and notifications are enabled.
   - Auto-apply is disabled, keeping the final action under administrative control.
   - A no-response decision is configured as **Approve**.

3. **Monitor and escalate approaching deadlines**
   - A separate certificate-authenticated monitoring script is intended to run from Windows Task Scheduler every **14 days**.
   - It evaluates access-review instances approaching their deadline.
   - `-DaysBeforeDeadline 14` provides a two-week monitoring window, so each scheduled execution can identify unreviewed items whose deadlines fall within the next 14 days.
   - Fallback administrators can be notified when reviews are nearing their deadline with outstanding/unreviewed members.

## Architecture

```text
                    Microsoft 365 Groups
                            |
                            v
              +---------------------------+
              | 1. Group Discovery        |
              | Get-M365GroupsAndOwners   |
              +-------------+-------------+
                            |
                            v
                 Group / Owner assessment
                            |
                            v
              +---------------------------+
              | 2. Access Review Creation |
              | New-M365GroupAccessReviews|
              +-------------+-------------+
                            |
                            v
             Recurring Entra Access Reviews
                    Every 6 months
                     14-day window
                            |
                            v
              +---------------------------+
              | 3. Scheduled Monitoring   |
              | Certificate Authentication|
              | Runs every 14 days        |
              +-------------+-------------+
                            |
                            v
                 Deadline / review scan
                            |
                            v
                    Admin escalation
```

## Repository contents

| Script | Purpose |
|---|---|
| `Get-M365GroupsAndOwners.ps1` | Discovers M365 Groups, owners, member counts, Team/SPO status, and owner account health; exports an Excel report. |
| `New-M365GroupAccessReviews.ps1` | Creates recurring Entra ID Access Reviews for M365 Groups from CSV input. |
| `Monitor-M365GroupAccessReviews.ps1` | Certificate-authenticated scheduled monitoring/escalation script. **Not included in the current public sample unless added separately.** |

### Generated files

The scripts can also generate:

- M365 Group inventory Excel reports
- Access Review creation CSV logs
- Runtime/reporting output

Do not commit generated reports or logs if they contain production identities or other tenant data.

## Prerequisites

### PowerShell modules

The discovery script uses:

- `Microsoft.Graph.Authentication`
- `Microsoft.Graph.Groups`
- `Microsoft.Graph.Users`
- `Microsoft.Graph.Sites`
- `ExchangeOnlineManagement`
- `ImportExcel`

The Access Review creation script uses:

- `Microsoft.Graph.Authentication`
- `Microsoft.Graph.Identity.Governance`

### Licensing

Access Review functionality requires the appropriate **Microsoft Entra ID Governance / Entra ID P2** licensing for the tenant and users covered by the implementation.

### Permissions

The discovery script requests:

```text
Group.Read.All
User.Read.All
Directory.Read.All
Sites.Read.All
```

The Access Review creation script requests:

```text
AccessReview.ReadWrite.All
Group.Read.All
User.Read.All
```

The certificate-authenticated monitoring script should be configured with the minimum Microsoft Graph **application permissions** required for its actual operations. If it sends mail through Microsoft Graph, `Mail.Send` is required.

> Use least privilege and grant only the permissions required by the deployment.

## Script 1 — Discover M365 Groups and owner health

```powershell
.\Get-M365GroupsAndOwners.ps1
```

Optional output path:

```powershell
.\Get-M365GroupsAndOwners.ps1 -OutputPath "C:\Reports\M365Groups.xlsx"
```

The resulting workbook contains:

- **M365 Groups** — complete inventory
- **Groups Without Owners** — groups requiring ownership remediation
- **Owners - Disabled or Shared** — groups with disabled/shared-mailbox owners

The script uses interactive authentication for Graph and Exchange Online.

## Script 2 — Create Access Reviews

Prepare a CSV containing:

```csv
GroupId,GroupDisplayName,Admins
00000000-0000-0000-0000-000000000000,Example Group,admins@example.com
```

`Admins` contains one or more **mail-enabled admin group email addresses**, separated by `;`.

Run:

```powershell
.\New-M365GroupAccessReviews.ps1 -CsvPath ".\M365Groups.csv"
```

The script protects against duplicate review definitions by checking for an existing review with the same display name.

### Review behavior

Each generated review is configured for:

- Target: M365 Group transitive membership
- Reviewers: current group owners
- Fallback reviewers: members of designated admin group(s)
- Recurrence: every 6 months
- Instance duration: 14 days
- Notifications: enabled
- Reminders: enabled
- Justification on approval: required
- Recommendations: enabled
- Auto-apply: disabled
- Default decision: Approve
- No end date

## Script 3 — Scheduled monitoring and escalation

The monitoring component is designed for Windows Task Scheduler and certificate-based application authentication.

Recommended execution model:

```text
Task Scheduler
      |
      | Every 14 days
      v
Certificate-authenticated PowerShell script
      |
      | -DaysBeforeDeadline 14
      v
Scan reviews approaching deadline
      |
      v
Find unreviewed items
      |
      v
Notify fallback administrators
```

The important scheduling relationship is:

> **Task frequency = 14 days**  
> **Deadline look-ahead = 14 days**

Using `-DaysBeforeDeadline 14` means the script evaluates reviews with deadlines in the next two weeks during each scheduled run. This is intended to provide coverage across the two-week execution window without requiring the task to run daily.

### Certificate authentication

For unattended execution, the application registration should use certificate-based authentication rather than an interactive sign-in.

Typical configuration parameters include:

```text
TenantId
ClientId
CertificateThumbprint
SenderMailbox
DaysBeforeDeadline = 14
```

Keep certificate private keys, tenant-specific configuration, mailbox information, and production identifiers outside source control.

## Security considerations

This project works with identity and access-governance data. Treat its output as sensitive operational data.

Do **not** commit:

- Tenant IDs if your organization's policy treats them as confidential
- Client/application IDs where policy requires them to remain private
- Certificate private keys
- Certificate passwords
- Certificate thumbprints if considered sensitive by your organization
- Production mailbox addresses
- Group IDs or production group names
- User UPNs
- Exported Excel reports
- Access Review logs containing production identities
- Secrets or tokens

Use a configuration file/template containing placeholders and inject production values securely.

Example:

```powershell
$TenantId = "<TENANT-ID>"
$ClientId = "<CLIENT-ID>"
$CertificateThumbprint = "<CERTIFICATE-THUMBPRINT>"
$SenderMailbox = "<SENDER-MAILBOX>"
```

## Important design note

This implementation deliberately keeps **auto-apply disabled**.

That means an Access Review can record reviewer decisions without automatically changing group membership. Organizations can therefore add their own approval/remediation process before membership changes are applied.

Also note that configuring a **default decision of Approve** means that a reviewer who does not respond within the configured review period can result in an approval decision. Review this behavior carefully against your organization's governance policy before deploying to production.

## Validation checklist

Before production deployment:

- [ ] Confirm Microsoft Entra ID Governance licensing.
- [ ] Confirm Graph permissions and admin consent.
- [ ] Test against a non-production M365 Group.
- [ ] Validate owner and fallback-reviewer behavior.
- [ ] Confirm the 14-day review duration.
- [ ] Confirm the 6-month recurrence.
- [ ] Confirm auto-apply remains disabled.
- [ ] Validate the no-response/default-decision policy.
- [ ] Configure certificate authentication for unattended monitoring.
- [ ] Test Task Scheduler execution under the intended service identity.
- [ ] Test escalation email delivery.
- [ ] Review generated logs for sensitive information.
- [ ] Add production secrets/configuration through a secure mechanism rather than source control.

## Future roadmap

The current implementation is intentionally focused on **M365 Groups**.

Potential future scope includes:

- Microsoft Teams
- SharePoint
- Enterprise Applications
- Service Principals
- Other Entra ID resources
- Additional review/remediation workflows
- Improved reporting and dashboards

More workloads will be added as the project evolves.

## Disclaimer

This repository is a reference/automation project and should be validated against your organization's security, identity-governance, compliance, and change-management requirements before production use.

Microsoft Graph APIs and Entra ID behavior can change over time. Always validate the scripts against your current tenant and module/API versions.

## Author

**rathoosh16**

GitHub:

`https://github.com/rathoosh16/M365-Access-Reviews-Automation`

---

If this project helps with your access-governance automation, feel free to ⭐ the repository and contribute improvements.
