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
This Script works for every NoSpamProxy version 11.0 and higher.
