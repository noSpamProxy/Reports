# Reports

This repository contains PowerShell Scripts to create advanced reports about your NoSpamProxy installation.

## Get-RejectionReport

This report displays information about all rejected e-mails and which filters and actions were involved. Use this version for NoSpamProxy 12.x and higher.

Currently, there are three versions of this script:

* [NoSpamProxy 12.x and later, access via Database](Get-RejectionReport%20(Database))
* [NoSpamProxy 12.x and later, access via WebServices](Get-RejectionReport%20(WebServices))
* [NoSpamProxy 11.x](11.x/Get-RejectionReport)

## Send-ReporttoUserswithAllBlockedEmails

Sends a report to every E-Mail address that contains all permanently blocked E-Mails in a specific period of time.

Currently, there are two version of this script:

* [NoSpamProxy 12.x and later](Send-ReporttoUserswithBlockedEmails)
* [NoSpamProxy 11.x](11.x/Send-ReporttoUsersWithBlockedEmails)

## Send-ReporttoUserswithBlockedEmailsWithAttachment

Sends a report to every E-Mail address that contains all permanently blocked E-Mails in a specific period of time that have been blocked because of an attachment.

* [NoSpamProxy 12.x and later](Send-ReporttoUserswithBlockedEmailsWithAttachment)
* [NoSpamProxy 11.x](11.x/Send-ReporttoUserswithBlockedEmailsWithAttachment)

## [Send-BlockReportToUsers](Send-BlockReportToUsers/Readme.md)

Rewritten version of "Send-ReporttoUserswithAllBlockedEmails" with enhanced features.

This script can be used to generate a report about E-Mails which where permanently blocked.  
It is possible to filter the results for a specific time duration and sends the report only to specific or all affected users.  
This script only uses the NoSpamProxy Powershell Cmdlets to generate the report file.



## [Get-DmarcReport](https://github.com/noSpamProxy/Reports/tree/master/Get-DMARCReport)

This report displays information about DMARC Check results for all successful inbound emails.

Currently, there are two versions of this script:

* [NoSpamProxy 12.2 and later](Get-DMARCReport)
* [NoSpamProxy 11.x](11.x/Get-DMARCReport)

## [Get-EncryptionReport](https://github.com/noSpamProxy/Reports/tree/master/Get-EncrxyptionReport)

This script generates a report over the count of encrypted and non encrypted mails.

## [Get-ExpiredTLSCertificates](https://github.com/noSpamProxy/Reports/tree/master/get-expiredTLSCertificates)

This report displays all expired certificates from e-mail servers that NoSpamProxy has sent outbound e-mails to.

## [get-certificatesInQuarantine](https://github.com/noSpamProxy/Reports/tree/master/get-certificatesInQuarantine)

This report displays a list of certificates in the certificate quarantine.

## [Get-NspMsgTrackAsCSV](https://github.com/noSpamProxy/Reports/tree/master/Get-NspMsgTrackAsCSV)

This report exports the message track as CSV file.


## [Get-NspTlsReport](https://github.com/noSpamProxy/Reports/tree/master/Get-NspTlsReport)

This report generates a HTML which contains an overview of your email communication using TLS.  
Requires NSP version 13.0.19147.917 or later.

## [Send-BlockReportToUsers](https://github.com/noSpamProxy/Reports/tree/master/Send-BlockReportToUsers)

This report generates a HTML which contains an overview of your email communication showing all permanently blocked emails.

## [Send-EncryptReportToUsers](https://github.com/noSpamProxy/Reports/tree/master/Send-EncryptReportToUsers)

This report generates a HTML which contains an overview of your email communication showing which emails where encrypted.  

