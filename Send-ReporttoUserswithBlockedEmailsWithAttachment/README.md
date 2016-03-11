# Send-ReporttoUserswithBlockedEmailsWithAttachment.ps1
Sends a report to every E-Mail address that contains all permanently blocked E-Mails in a specific period of time that have been blocked because of an attachment. The report contains:
 - DeliveryStartTime
 - Sender address
 - Subject
 - Name of blocked file


###Usage
`Send-ReporttoUserswithBlockedEmailsWithAttachment -SMTPHost [-NumberOfDaysToReport] [-ReportSender] [-ReportSubject]`

- SMTPHost
```
Mandatory. Specifies the SMTP Host which will be used to send the email.
```
- NumberOfDaysToReport
```
Optional. Specifies the Number of days to report. Default value is "1".
```
- ReportSender
```
Optional. Specifies the Sender of the email. Default value is "NoSpamProxy Report Sender <nospamproxy@example.com>".
```
- ReportSubject
```
Optional. Specifies the Subject of the email. Default value is "Auswertung".
```
###Example
`.\Send-ReporttoUserswithBlockedEmailsWithAttachment.ps1 -SMTPHost mail.example.com`

###Supported NoSpamProxy Versions
This Script works for every NoSpamProxy version 10.x and higher.
