# Get-DMARCReport.ps1
Sends a report to the specified E-Mail address which contains the following information for every successful inbound email:
 - Domain information from envelope
 - Domain information from Header
 - SPF Check result
 - SPF Alignment result
 - DKIM Check result
 - DKIM Alignment result
 - DMARC Check Result
 - DMARC Policy (Applicable DMARC policy)
 - What if (Effective policy)
 - Date
 - Message ID 


###Authentication
The script signs in with `Connect-Nsp` and supports BASIC or integrated Windows (**NTLM**) authentication.

- **Credential**: Optional `PSCredential`. When set, BASIC authentication is used.
- **CredentialFileLocation**: Optional path to a saved BASIC credential file. The default is `NspBasicAuthCredential.bin` beside the script. Relative paths such as `./auth.bin` are resolved beside the script, not from the current PowerShell directory.
- **SetCredential**: Creates or replaces the credential file and exits. It is a standalone parameter set, so no report parameters are required.
- **UseIntegratedAuth**: Forces NTLM and ignores both `Credential` and the saved credential file.

Credential selection order is: `UseIntegratedAuth` (NTLM), explicit `Credential` (BASIC), saved credential file (BASIC), then NTLM when no credential is available. The saved file contains the username and a password encrypted with Windows DPAPI `LocalMachine`; it can be read by users on the same computer who have filesystem permission to read the file.

Create the default credential file:
```ps
.\Get-DMARCReport.ps1 -SetCredential
```

Create a credential file at a custom script-relative path and use it:
```ps
.\Get-DMARCReport.ps1 -SetCredential -CredentialFileLocation .\auth.bin
.\Get-DMARCReport.ps1 -CredentialFileLocation .\auth.bin
```

Hardcoded credential (for automation):
```ps
$pass = ConvertTo-SecureString 'PASSWORD' -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential('DOMAIN\\user', $pass)
.\Get-DMARCReport.ps1 -Credential $cred
```

###Usage
`Get-DMARCReport -SMTPHost -ReportRecipient [-ReportSubject] [-NumberOfDaysToReport] [-ReportSender]`

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
- NumberOfDaysToReport
```
Optional. Specifies the Number of days to report. Default value is "7".
```
- ReportSender
```
Optional. Specifies the Sender of the email. Default value is "NoSpamProxy Report Sender <nospamproxy@example.com>".
```

###Example
`.\Get-RejectionReport.ps1 -SMTPHost mail.example.com -ReportRecipient admin@example.com`

###Supported NoSpamProxy Versions
This Script works for every NoSpamProxy version 12.2 and higher.
