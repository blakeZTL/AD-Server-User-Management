# AD Group Membership Sync Script

This PowerShell script synchronizes an on-prem Active Directory security group with a filtered Excel list.

It:

- Loads users from an Excel spreadsheet
- Filters by **Job Series = 2152**
- Treats that filtered list as the **source of truth**
- Adds users missing from the group
- Removes users not in the Excel list
- Skips users already correctly assigned
- Outputs a detailed console summary and transcript log

---

## Requirements

### 1. ImportExcel Module

Used to read `.xlsx` files directly without requiring Excel to be installed.

Install:

```powershell
Install-Module ImportExcel -Scope CurrentUser
```

Documentation:

- GitHub: https://github.com/dfinke/ImportExcel
- PowerShell Gallery: https://www.powershellgallery.com/packages/ImportExcel

---

### 2. ActiveDirectory Module

Used for interacting with on-prem AD.

This module provides:

- `Get-ADUser`
- `Get-ADGroupMember`
- `Add-ADGroupMember`
- `Remove-ADGroupMember`

Install via RSAT (Windows 10/11):

```powershell
Get-WindowsCapability -Name RSAT.ActiveDirectory* -Online
Add-WindowsCapability -Name RSAT.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -Online
```

Documentation:

- ActiveDirectory Module  
  https://learn.microsoft.com/en-us/powershell/module/activedirectory/

- Get-ADUser  
  https://learn.microsoft.com/en-us/powershell/module/activedirectory/get-aduser

- Get-ADGroupMember  
  https://learn.microsoft.com/en-us/powershell/module/activedirectory/get-adgroupmember

- Add-ADGroupMember  
  https://learn.microsoft.com/en-us/powershell/module/activedirectory/add-adgroupmember

- Remove-ADGroupMember  
  https://learn.microsoft.com/en-us/powershell/module/activedirectory/remove-adgroupmember

---

## Expected Excel Format

The spreadsheet must contain the following columns:

| Column Name | Required | Description |
|------------|----------|-------------|
| Email | Yes | User email address |
| Job Series | Yes | Must equal `2152` to be included |

Only rows where **Job Series = 2152** will be processed.

Example:

| Email | Job Series |
|-------|-----------|
| james.b.bradford@faa.gov | 2152 |

---

## Configuration

Edit these variables at the top of the script:

```powershell
$group               = "V-AFG001-APP-Part48-GASA-Access"
$emailColumnName     = "Email"
$jobSeriesColumnName = "Job Series"
$filePath            = ".\Emails.xlsx"
$sheetName           = "Sheet1"
$targetJobSeries     = "2152"
$applyChanges        = $false
```

---

## Safety Mode (Dry Run)

By default:

```powershell
$applyChanges = $false
```

This means:

- No users are actually added
- No users are actually removed
- The script only reports what would happen

To apply changes:

```powershell
$applyChanges = $true
```

---

## How It Works

1. Load Excel file  
2. Filter rows where Job Series = 2152  
3. Resolve emails to AD users  
4. Build DistinguishedName sets for:
   - Target users (Excel)
   - Current group members (AD)
5. Compute delta:
   - **Add** = In Excel but not in group
   - **Remove** = In group but not in Excel
   - **Skip** = Already correct
6. Apply changes (if enabled)
7. Output summary and transcript log

---

## Output

The script provides:

- Real-time progress bars
- Timestamped console logs
- Add / Remove counts
- Not Found report
- Full transcript file

A transcript file is automatically created:

```
GroupSync-<GroupName>-YYYYMMDD-HHMMSS.log
```

---

## Performance Design

Optimized for large groups (1500+ members):

- Avoids per-user AD queries where possible
- Uses DistinguishedName comparisons
- Uses HashSets for fast membership checks (O(1))
- Throttled logging to prevent console slowdown

The script scales well as the group grows.

---

## Operational Considerations

- Requires AD write permissions to modify group membership
- Recursive group membership is supported
- Only user objects are evaluated (nested groups are ignored)
- Email is matched against:
  - `mail`
  - `UserPrincipalName` (fallback)

---

## Recommended Workflow

1. Run script with `$applyChanges = $false`
2. Review transcript log
3. Validate add/remove counts
4. Set `$applyChanges = $true`
5. Re-run and confirm results

---

## Common Issues

| Issue | Cause |
|-------|-------|
| Users not found | Email does not match `mail` or `UserPrincipalName` in AD |
| Access denied | Insufficient permissions to modify the group |
| Slow execution | Domain controller latency or replication delays |

---

## Notes

Designed for enterprise AD environments where:

- Excel serves as the authoritative roster
- AD group must reflect filtered role membership
- Auditability and logging are required

Use in accordance with your organization’s AD change control policies.