# Get-RejectionReport.ps1
Sends a report to the specified E-Mail address which contains the following information:
 - Total count of rejected emails
 - Total count of permanently rejected emails
 - Total count of temporary rejected emails
 - Total count of rejected emails on envelope level
 - Total count of rejected emails on body level
 - Total count of rejected emails, sorted by filters and actions.

###Usage
`Get-RejectionReport -SMTPHost -ReportRecipient [-ReportSubject] [-NumberOfDaysToReport] [-ReportSender]`

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
Optional. Specifies the Subject of the email. Default value is "Auswertung".
```
- NumberOfDaysToReport
```
Optional. Specifies the Number of days to report. Default value is "1".
```
- ReportSender
```
Optional. Specifies the Sender of the email. Default value is "NoSpamProxy Report Sender <nospamproxy@example.com>".
```

###Example
`.\Get-RejectionReport.ps1 -SMTPHost mail.example.com -ReportRecipient admin@example.com`

###Supported NoSpamProxy Versions
This Script works for every NoSpamProxy version 10.x and higher.
