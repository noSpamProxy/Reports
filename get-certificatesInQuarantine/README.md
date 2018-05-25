# get-certificatesInQuarantine.ps1
Sends a report to the specified E-Mail address which contains a list of certificates in the certificate quarantine. The report contains the following information:
 - SubjectDistinguishedName
 - IssuerDistinguishedName
 - ValidTo


###Usage
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

###Example
`.\get-certificatesInQuarantine.ps1 -SMTPHost mail.example.com -ReportRecipient admin@example.com`

###Supported NoSpamProxy Versions
This Script works for every NoSpamProxy version 12.1 and higher.
