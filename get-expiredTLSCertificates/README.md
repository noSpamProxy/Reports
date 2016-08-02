# get-expiredTLSCertificates.ps1
Sends a report to the specified E-Mail address which contains the following information:
 - Recipient Domain
 - Expiration Date of the certificate
 - IP-Address of the Target Server
 - Certificate Subject of the expired certificate

###Usage
`get-expiredTLSCertificates -SMTPHost -ReportRecipient [-ReportSubject] [-NumberOfDaysToReport] [-ReportSender]`

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
Optional. Specifies the Subject of the email. Default value is "Auswertung der abgelaufenen TLS-Zertifikate".
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
`.\get-expiredTLSCertificates.ps1 -SMTPHost mail.example.com -ReportRecipient admin@example.com`

###Supported NoSpamProxy Versions
This Script works for every NoSpamProxy version 11.x and higher.
