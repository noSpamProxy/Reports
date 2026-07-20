# Get-EncryptionReport.ps1
This script generates a report over the count of encrypted and non encrypted mails. It seperates results in both directions and timespan (last 30 days and 30 days before that timespan).  
Report will be send as HTML attachment to a provided email address.

## Authentication
The script signs in with `Connect-Nsp` and supports BASIC or integrated Windows (**NTLM**) authentication.

- **Credential**: Optional `PSCredential`. When set, BASIC authentication is used.
- **CredentialFileLocation**: Optional path to a saved BASIC credential file. The default is `NspBasicAuthCredential.bin` beside the script. Relative paths such as `./auth.bin` are resolved beside the script, not from the current PowerShell directory.
- **SetCredential**: Creates or replaces the credential file and exits. It is a standalone parameter set, so no report parameters are required.
- **UseIntegratedAuth**: Forces NTLM and ignores both `Credential` and the saved credential file.

Credential selection order is: `UseIntegratedAuth` (NTLM), explicit `Credential` (BASIC), saved credential file (BASIC), then NTLM when no credential is available. The saved file contains the username and a password encrypted with Windows DPAPI `LocalMachine`; it can be read by users on the same computer who have filesystem permission to read the file.

Create the default credential file:
```ps
.\Get-EncryptionReport.ps1 -SetCredential
```

Create a credential file at a custom script-relative path and use it:
```ps
.\Get-EncryptionReport.ps1 -SetCredential -CredentialFileLocation .\auth.bin
.\Get-EncryptionReport.ps1 -CredentialFileLocation .\auth.bin
```

Hardcoded credential (for automation):
```ps
$pass = ConvertTo-SecureString 'PASSWORD' -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential('DOMAIN\\user', $pass)
.\Get-EncryptionReport.ps1 -Credential $cred
```

## Usage  
```ps 
Get-EncryptionReport.ps1 [-smtphost <String>] [-ReportSender <String>] [-ReportRecipient <String>]
```

## Parameters
### Credential
Optional `PSCredential` for NoSpamProxy BASIC authentication. It takes precedence over a saved credential file.

### CredentialFileLocation
Optional path to the saved NoSpamProxy BASIC credential file. The default is `NspBasicAuthCredential.bin` in the script directory; relative paths are also resolved from the script directory.

### SetCredential
Creates or replaces the saved NoSpamProxy credential file and then exits. Prompts for the NoSpamProxy username and password; no report parameters are required.

### UseIntegratedAuth
Forces integrated Windows (NTLM) authentication and ignores `Credential` and `CredentialFileLocation`.

### smtphost  
Specifies the Hostname or IP of the mailserver the report will be sent to.

### ReportSender
Specifies the mailaddress used as sender for the report.

### ReportRecipient
Specifies the recipient address for the report.

## Examples
### Example 1
```ps
.\Get-EncryptionReport.ps1 -smtphost "mailserver.company" -ReportSender "mailgateway@company" -ReportRecipient "admin@company"
```