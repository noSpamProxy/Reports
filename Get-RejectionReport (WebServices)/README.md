# Get-RejectionReport.ps1

Sends a report to the specified E-Mail address which contains the following information:

- Total count of processed emails.
- Total count of sent emails.
- Total count of received emails.
- Total count of rejected emails.
- Total count of rejected emails for each Filter and Action.

This version uses the NoSPamProxy WebServices to acquire the required information. This is suited for small to medium sized installations. For large installations, the [Database Version](../Get-RejectionReport%20(Database)) version instead.

## Authentication
The script signs in with `Connect-Nsp` and supports BASIC or integrated Windows (**NTLM**) authentication.

- **Credential**: Optional `PSCredential`. When set, BASIC authentication is used.
- **CredentialFileLocation**: Optional path to a saved BASIC credential file. The default is `NspBasicAuthCredential.bin` beside the script. Relative paths such as `./auth.bin` are resolved beside the script, not from the current PowerShell directory.
- **SetCredential**: Creates or replaces the credential file and exits. It is a standalone parameter set, so no report parameters are required.
- **UseIntegratedAuth**: Forces NTLM and ignores both `Credential` and the saved credential file.

Credential selection order is: `UseIntegratedAuth` (NTLM), explicit `Credential` (BASIC), saved credential file (BASIC), then NTLM when no credential is available. The saved file contains the username and a password encrypted with Windows DPAPI `LocalMachine`; it can be read by users on the same computer who have filesystem permission to read the file.

Create the default credential file:
```ps
.\Get-RejectionReport.ps1 -SetCredential
```

Create a credential file at a custom script-relative path and use it:
```ps
.\Get-RejectionReport.ps1 -SetCredential -CredentialFileLocation .\auth.bin
.\Get-RejectionReport.ps1 -CredentialFileLocation .\auth.bin
```

Hardcoded credential (for automation):
```ps
$pass = ConvertTo-SecureString 'PASSWORD' -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential('DOMAIN\\user', $pass)
.\Get-RejectionReport.ps1 -Credential $cred
```

## Usage

```ps
Get-RejectionReport -SMTPHost -ReportRecipient [-ReportSubject] [-NumberOfDaysToReport] [-ReportSender]`
```

- **SMTPHost**: Mandatory. Specifies the SMTP Host which will be used to send the email.
- **ReportRecipient**: Mandatory. Specifies the Recipient of the email.
- **ReportSubject**: Optional. Specifies the Subject of the email. Default value is "Auswertung".
- **ReportSender**: Mandatory. Specifies the Sender of the email.
- **NumberOfDaysToReport**: Optional. Specifies the Number of days to report. Default value is "7".
- **TopAddressesCount**: Optional. How many Addresses to include in the top Senders/Recipients/Spammers. Defaults to 5
- **ExcludeFromTopAddresses**: Optional. Specify addresses to ignore when evaluating top Senders/Recipients. Useful for dropping obvious addresses from the list to get a more informative statistic

## Example

```ps
.\Get-RejectionReport.ps1 -SMTPHost mail.example.com -ReportRecipient admin@example.com` -ReportSender "NoSpamProxy Report Sender <nospamproxy@example.com>"
```

## Supported NoSpamProxy Versions

This Script works for NoSpamProxy version 12.2 and higher.
