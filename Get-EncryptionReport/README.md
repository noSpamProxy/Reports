# Get-EncryptionReport.ps1
This script generates a report over the count of encrypted and non encrypted mails. It seperates results in both directions and timespan (last 30 days and 30 days before that timespan).  
Report will be send as HTML attachment to a provided email address.

## Usage  
```ps 
Get-EncryptionReport.ps1 [-smtphost <String>] [-ReportSender <String>] [-ReportRecipient <String>]
```

## Parameters
### smtphost  
Specifies the Hostname or IP of the mailserver the report will be sent to.

### ReportSender
Specifies the mailaddress used as sender for the report.

### ReportRecipient
Specifies the recipient address for the report.

## Examples
### Example 1
```ps
.\Get-EncryptionReport.ps1 -smtphost "mailserver.company" -ReportSender "mailgateway@company" -ReportRecipient "admin@company"
```