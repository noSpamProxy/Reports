# get-certificatesInQuarantine.ps1
Sends a report to the specified E-Mail address which contains a list of certificates in the certificate quarantine. If no certificates are in the quarantine, no report is sent. The report contains the following information:
 - SubjectDistinguishedName
 - IssuerDistinguishedName
 - ValidTo


### Authentication
The script signs in with `Connect-Nsp` and supports BASIC or integrated Windows (**NTLM**) authentication.

- **Credential**: Optional `PSCredential`. When set, BASIC authentication is used.
- **CredentialFileLocation**: Optional path to a saved BASIC credential file. The default is `NspBasicAuthCredential.bin` beside the script. Relative paths such as `./auth.bin` are resolved beside the script, not from the current PowerShell directory.
- **SetCredential**: Creates or replaces the credential file and exits. It is a standalone parameter set, so no report parameters are required.
- **UseIntegratedAuth**: Forces NTLM and ignores both `Credential` and the saved credential file.

Credential selection order is: `UseIntegratedAuth` (NTLM), explicit `Credential` (BASIC), saved credential file (BASIC), then NTLM when no credential is available. The saved file contains the username and a password encrypted with Windows DPAPI `LocalMachine`; it can be read by users on the same computer who have filesystem permission to read the file.

Create the default credential file:
```ps
.\get-certificatesInQuarantine.ps1 -SetCredential
```

Create a credential file at a custom script-relative path and use it:
```ps
.\get-certificatesInQuarantine.ps1 -SetCredential -CredentialFileLocation .\auth.bin
.\get-certificatesInQuarantine.ps1 -CredentialFileLocation .\auth.bin
```

Hardcoded credential (for automation):
```ps
$pass = ConvertTo-SecureString 'PASSWORD' -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential('DOMAIN\\user', $pass)
.\get-certificatesInQuarantine.ps1 -Credential $cred
```

## Usage
`get-certificatesInQuarantine.ps1 -SMTPHost -ReportRecipient [-ReportSubject] [-ReportSender]`

- SMTPHost
```
Mandatory. Specifies the SMTP Host which will be used to send the email.
```
- ReportRecipient
```
Mandatory. Specifies the Recipient of the email.
```
- ReportSubject
```
Optional. Specifies the Subject of the email. Default value is "DMARC Auswertung der letzten Woche".
```
- ReportSender
```
Optional. Specifies the Sender of the email. Default value is "NoSpamProxy Report Sender <nospamproxy@example.com>".
```

### Example
`.\get-certificatesInQuarantine.ps1 -SMTPHost mail.example.com -ReportRecipient admin@example.com`

### Supported NoSpamProxy Versions
This Script works for every NoSpamProxy version 12.1 and higher.
