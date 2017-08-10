# Get-RejectionReport.ps1

Sends a report to the specified E-Mail address which contains the following information:

- Total count of processed emails.
- Total count of sent emails.
- Total count of received emails.
- Total count of rejected emails.
- Total count of rejected emails for each Filter and Action.

This version accesses the NoSpamProxy database directly and is much faster than the other version. The drawback is that most NoSpamProxy installations do not allow remote access to the Database so this script needs to be run locally on the Intranet Role.

## Usage

```ps
Get-RejectionReport -SMTPHost -ReportRecipient [-ReportSubject] [-NumberOfDaysToReport] [-ReportSender]`
```

- **SMTPHost**: Mandatory. Specifies the SMTP Host which will be used to send the email.
- **ReportRecipient**: Mandatory. Specifies the Recipient of the email.
- **ReportSubject**: Optional. Specifies the Subject of the email. Default value is "Auswertung".
- **NumberOfDaysToReport**: Optional. Specifies the Number of days to report. Default value is "7".
- **ReportSender**: Optional. Specifies the Sender of the email. Default value is "NoSpamProxy Report Sender <nospamproxy@example.com>".
- **SqlServer**: Optional. The name of the Database server (including instance name, if any). Defaults to (local)\NoSpamProxyDB.
- **Database**: Optional. Name the Database to query. Defaults to "NoSpamProxyAddressSynchronization".
- **Credential**: Optional. Username and password for the SQL authentication on the database server. If not set, Integrated authentication is used, which is the default.


## Example

```ps
.\Get-RejectionReport.ps1 -SMTPHost mail.example.com -ReportRecipient admin@example.com'
```

## Supported NoSpamProxy Versions

This Script works for NoSpamProxy version 12.x and higher.