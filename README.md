# CMDB to Intune Device Tiering

PowerShell automation that imports a CMDB CSV, maps devices to Intune managed
devices through Microsoft Graph, and assigns Intune Device Categories such as
`Tier-0` through `Tier-5`.

The script includes:

- tier normalisation (`0`, `Tier 0`, `Tier-0`, `T0`, and equivalents);
- matching by device name, with serial number as a fallback;
- duplicate-match protection;
- optional creation of missing Intune Device Categories;
- `-WhatIf` / `ShouldProcess` support;
- a detailed CSV execution report.

## Requirements

- Microsoft Graph PowerShell SDK.
- Graph delegated permissions:
  - `DeviceManagementManagedDevices.ReadWrite.All`
  - `DeviceManagementConfiguration.ReadWrite.All`
- A CSV containing `Name` and `Tier`; `Serial Number` is optional.

## Example

```powershell
.\Set-IntuneDeviceTierFromCMDB.ps1 `
    -CsvPath .\examples\cmdb-servers.example.csv `
    -WhatIf
```

After reviewing the report, rerun without `-WhatIf`. Use
`-CreateMissingCategories` only when category creation is intended.

## Security

No tenant IDs, device names, serial numbers, credentials, or organisation
names are embedded in the repository. The sample CSV contains synthetic data.
