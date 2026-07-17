# Install AWS CLI v2
RUN apt-get update \
	&& apt-get install -y --no-install-recommends \
		ca-certificates \
		curl \
		unzip 
# Get-NspLicenseReport.ps1

Generates a report for the usage of each module containing the following information:

 - Total Users who used the module
 - List of each user who used the module
 
It is also possible to generate a domain based license report, please the alternative script therefore:  
[Get-NspLicenseReport_(Database)](https://github.com/noSpamProxy/Reports/tree/master/Get-NspLicenseReport_(Database))


## Authentication
The script signs in with `Connect-Nsp` and supports BASIC or integrated Windows (**NTLM**) authentication.

- **Credential**: Optional `PSCredential`. When set, BASIC authentication is used.
- **CredentialFileLocation**: Optional path to a saved BASIC credential file. The default is `NspBasicAuthCredential.bin` beside the script. Relative paths such as `./auth.bin` are resolved beside the script, not from the current PowerShell directory.
- **SetCredential**: Creates or replaces the credential file and exits. It is a standalone parameter set, so no report parameters are required.
- **UseIntegratedAuth**: Forces NTLM and ignores both `Credential` and the saved credential file.

Credential selection order is: `UseIntegratedAuth` (NTLM), explicit `Credential` (BASIC), saved credential file (BASIC), then NTLM when no credential is available. The saved file contains the username and a password encrypted with Windows DPAPI `LocalMachine`; it can be read by users on the same computer who have filesystem permission to read the file.

Create the default credential file:
```ps
.\Get-NspLicenseReport.ps1 -SetCredential
```

Create a credential file at a custom script-relative path and use it:
```ps
.\Get-NspLicenseReport.ps1 -SetCredential -CredentialFileLocation .\auth.bin
.\Get-NspLicenseReport.ps1 -ExportFilePath "C:\Reports\Usage.txt" -CredentialFileLocation .\auth.bin
```

Hardcoded credential (for automation):
```ps
$pass = ConvertTo-SecureString 'PASSWORD' -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential('DOMAIN\\user', $pass)
.\Get-NspLicenseReport.ps1 -Credential $cred
```

## Usage 

```ps
Get-NspLicenseReport -ExportFilePath "PATH_TO_EXPORT_FILE"
```

- **ExportFilePath**: Mandatory. Specifies the path where the report will be saved.

## Example

```ps
.\Get-NspLicenseReport.ps1 -ExportFilePath "C:\example\Usage.txt"
```

## Supported NoSpamProxy Versions
This script works for NoSpamProxy version 13 or higher. It depends on PowerShell version 4 (will be checked automatically).
