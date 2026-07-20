# Send-ReporttoUserswithBlockedEmailsWithAttachment.ps1

Sends a report to every E-Mail address that contains all permanently blocked E-Mails in a specific period of time that have been blocked because of an attachment. The report contains:

- DeliveryStartTime
- Sender address
- Subject
- Name of blocked file

## Authentication
The script signs in with `Connect-Nsp` and supports BASIC or integrated Windows (**NTLM**) authentication.

- **Credential**: Optional `PSCredential`. When set, BASIC authentication is used.
- **CredentialFileLocation**: Optional path to a saved BASIC credential file. The default is `NspBasicAuthCredential.bin` beside the script. Relative paths such as `./auth.bin` are resolved beside the script, not from the current PowerShell directory.
- **SetCredential**: Creates or replaces the credential file and exits. It is a standalone parameter set, so no report parameters are required.
- **UseIntegratedAuth**: Forces NTLM and ignores both `Credential` and the saved credential file.

Credential selection order is: `UseIntegratedAuth` (NTLM), explicit `Credential` (BASIC), saved credential file (BASIC), then NTLM when no credential is available. The saved file contains the username and a password encrypted with Windows DPAPI `LocalMachine`; it can be read by users on the same computer who have filesystem permission to read the file.

Create the default credential file:
```ps
.\Send-ReporttoUserswithBlockedEmailsWithAttachment.ps1 -SetCredential
```

Create a credential file at a custom script-relative path and use it:
```ps
.\Send-ReporttoUserswithBlockedEmailsWithAttachment.ps1 -SetCredential -CredentialFileLocation .\auth.bin
.\Send-ReporttoUserswithBlockedEmailsWithAttachment.ps1 -SMTPHost mail.example.com -CredentialFileLocation .\auth.bin
```

Hardcoded credential (for automation):
```ps
$pass = ConvertTo-SecureString 'PASSWORD' -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential('DOMAIN\\user', $pass)
.\Send-ReporttoUserswithBlockedEmailsWithAttachment.ps1 -Credential $cred
```

## Usage

```ps
Send-ReporttoUserswithBlockedEmailsWithAttachment -SMTPHost [-NumberOfDaysToReport] [-ReportSender] [-ReportSubject]
```

- **SMTPHost**: Mandatory. Specifies the SMTP Host which will be used to send the email.
- **NumberOfDaysToReport**: Optional. Specifies the Number of days to report. Default value is "1".
- **ReportSender**: Optional. Specifies the Sender of the email. Default value is "NoSpamProxy Report Sender <nospamproxy@example.com>".
- **ReportSubject**: Optional. Specifies the Subject of the email. Default value is "Auswertung".
- **TenantPrimaryDomain**: Necessary to use if the provider mode for v14 is enabled. It specifies the desired tenant for the runtime environment.

## Example

```ps
.\Send-ReporttoUserswithBlockedEmailsWithAttachment.ps1 -SMTPHost mail.example.com`
```

## Supported NoSpamProxy Versions

This Script works for NoSpamProxy 12.x or later. 
