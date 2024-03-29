# Send-ReporttoUserswithAllBlockedEmails.ps1

Sends a report to every E-Mail address that contains all permanently blocked E-Mails in a specific period of time. The report contains:

- DeliveryStartTime
- Sender address
- Recipient address
- Subject

## Usage

```ps
Send-ReporttoUserswithAllBlockedEmails -SMTPHost [-NumberOfDaysToReport] [-ReportSender] [-ReportSubject]`
```

- **SMTPHost**: Mandatory. Specifies the SMTP Host which will be used to send the email.
- **NumberOfDaysToReport**: Optional. Specifies the Number of days to report. Default value is "1".
- **ReportSender**: Optional. Specifies the Sender of the email. Default value is "NoSpamProxy Report Sender <nospamproxy@example.com>".
- **ReportSubject**: Optional. Specifies the Subject of the email. Default value is "Auswertung".
- **TenantPrimaryDomain**: Necessary to use if the provider mode for v14 is enabled. It specifies the desired tenant for the runtime environment.

## Example

```ps
.\Send-ReporttoUserswithAllBlockedEmails.ps1 -SMTPHost mail.example.com`
```

## Supported NoSpamProxy Versions

This Script works with NoSpamProxy 12.2 and higher.
