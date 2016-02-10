# get-rejectionreport.ps1
Sends a report to the specified E-Mail address which contains the following information:
 - Total count of rejected emails
 - Total count of permanently rejected emails
 - Total count of temporary rejected emails
 - Total count of rejected emails on envelope level
 - Total count of rejected emails on body level
 - Total count of rejected emails, sorted by filters and actions.

###Usage
`get-rejectionreport -<NumberOfDaysToReport> -<SMTPHost> -<ReportSender> -<ReportRecipient> -<ReportSubject>`

- NumberOfDaysToReport
```
Is not mandatory. Specifies the Number of days to report. Default value is "1".
```

- SMTPHost
```
Is mandatory. Specifies the SMTP Host which will be used to send the email.
```

- ReportSender
```
Is not mandatory. Specifies the Sender of the email. Default value is "NoSpamProxy Report Sender <nospamproxy@example.com>".
```
- ReportRecipient
```
Is mandatory. Specifies the Recipient of the email.
```
- ReportSubject
```
Is not mandatory. Specifies the Subject of the email. Default value is "Auswertung".
```

###Supported NoSpamProxy Versions
This Script works for every NoSpamProxy version 10.x and higher.
