# Reports

This repository contains PowerShell Scripts to create advanced reports about your NoSpamProxy installation.

## Get-RejectionReport

This report displays information about all rejected e-mails and which filters and actions were involved. Use this version for NoSpamProxy 12.x and higher.

Currently, there are three versions of this script:

* [NoSpamProxy 12.x and later, access via Database](Get-RejectionReport%20(Database))
* [NoSpamProxy 12.x and later, access via WebServices](Get-RejectionReport%20(WebServices))
* [NoSpamProxy 11.x](11.x/Get-RejectionReport)

## [Send-ReporttoUserswithAllBlockedEmails](https://github.com/noSpamProxy/Reports/tree/master/Send-ReporttoUserswithAllBlockedEmails)

Sends a report to every E-Mail address that contains all permanently blocked E-Mails in a specific period of time.

Currently, there are two version of this script:

* [NoSpamProxy 12.x and later](Send-ReporttoUserswithBlockedEmails)
* [NoSpamProxy 11.x](11.x/Send-ReporttoUserswithBlockedEmails)

## [Send-ReporttoUserswithBlockedEmailsWithAttachment](https://github.com/noSpamProxy/Reports/tree/master/Send-ReporttoUserswithBlockedEmailsWithAttachment)

Sends a report to every E-Mail address that contains all permanently blocked E-Mails in a specific period of time that have been blocked because of an attachment.

## [Get-DmarcReport](https://github.com/noSpamProxy/Reports/tree/master/get-DMARCReport)

This report displays information about DMARC Check results for all successful inbound emails.

## [Get-ExpiredTLSCertificates](https://github.com/noSpamProxy/Reports/tree/master/get-expiredTLSCertificates)

This report displays all expired certificates from e-mail servers that NoSpamProxy has sent outbound e-mails to.
