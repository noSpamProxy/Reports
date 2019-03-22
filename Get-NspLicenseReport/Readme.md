# Get-NspLicenseReport.ps1

Generates a report for the usage of each module containing the following information:

 - Total Users who used the module
 - List of each user who used the module

## Usage 

```ps
Get-NspLicenseReport -ExportFilePath
```

- **ExportFilePath**: Mandatory. Specifies the path where the report will be saved.

## Example

```ps
.\Get-NspLicenseReport.ps1 -ExportFilePath Usage.txt
```

## Supported NoSpamProxy Versions
This script works for NoSpamProxy version 13 or higher. It depends on PowerShell version 4 (will be checked automatically).