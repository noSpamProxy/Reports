Param
(
[Parameter(Mandatory=$true)][string] $SMTPHost,
[Parameter(Mandatory=$false)][int] $NumberOfDaysToReport = 1,
[Parameter(Mandatory=$false)][string] $ReportSender = "NoSpamProxy Report Sender <nospamproxy@example.com>",
[Parameter(Mandatory=$false)][string] $ReportSubject = "Auswertung der abgewiesenen E-Mails an Sie"
)

$dateStart = (Get-Date).AddDays(-$NumberOfDaysToReport)
$reportaddressesFileName = $Env:TEMP + "\reportaddresses.txt"
$reportAddresses = $null
$existingAddresses = $null

Write-Host "Getting MessageTrackInformation.."
$messageTracks = Get-NSPMessageTrack -Status PermanentlyBlocked -From $dateStart -Directions FromExternal |Get-NSPMessageTrackdetails
Write-Host "Done."
Write-Host "Create Reportaddresses-File"
[String[]] $existingAddresses = @()

foreach ($messageTrack in $messageTracks)
{
$messageRecipients = $messageTrack.Recipients
foreach ($messageRecipient in $messageRecipients)
	{
		$messageRecipientAddress = ($messageRecipient.Address).trim()
		if ($existingAddresses -notcontains $messageRecipientAddress) {
		$existingAddresses = $existingAddresses + $messageRecipientAddress
		}
	}
}
Set-Content $reportaddressesFileName $existingAddresses
Write-Host "Done."
Write-Host "Generating and sending reports for the following e-mail addresses:"

$existingAddresses | ForEach-Object {

Write-Host $_
$dateStart = (Get-Date).AddDays(-$NumberOfDaysToReport)
$reportFileName = $Env:TEMP + "\reject-analysis.html"

$htmlbody1 ="<html>
		<head>
			<title>Abgewiesene E-Mails an Sie</title>
			<style>
				  table, td, th { border: 1px solid black; border-collapse: collapse; }
				#headerzeile         {background-color: #DDDDDD;}
			</style>
		</head>
	<body style=font-family:arial>
		<h1>Abgewiesene E-Mails an Sie</h1>
		<br>
		<table>
			<tr id=headerzeile>
				<td><h3>Uhrzeit</h3></td><td><h3>Absender</h3></td><td><h3>Empf&auml;ngeradresse</h3></td><td><h3>Betreff</h3></td>
			</tr>
			"

$MTracks = Get-NSPMessageTrack -Between1 $_ -Status PermanentlyBlocked -From $dateStart -Directions FromExternal |Get-NSPMessageTrackdetails
$htmlbody2 =@()
foreach ($validationItem in $MTracks) 
{
$NSPStartTime = $validationItem.DeliveryStartTime
$NSPSender = $validationItem.Sender
$NSPSubject = $validationItem.Subject
$htmlbody2 +=("<tr><td width=150px>" +$NSPStartTime + "</td><td>" +$NSPSender +"</td><td>" +$_ + "</td><td>" +$NSPSubject + "</td></tr>")	
}

$htmlbody3="</table>
	</body>
	</html>"
$htmlout=$htmlbody1+$htmlbody2+$htmlbody3

$htmlout | Out-File $reportFileName
Send-MailMessage -SmtpServer $SmtpHost -From $ReportSender -To $_ -Subject $ReportSubject -Body "Im Anhang dieser E-Mail finden Sie den Bericht mit der Auswertung der abgewiesenen E-Mails aufgrund von Anh&auml;ngen an der E-Mail." -Attachments $reportFileName
Remove-Item $reportFileName
}
Write-Host "Done."
Write-Host "Doing some cleanup...."
Remove-Item $reportaddressesFileName
Write-Host "Done."
# SIG # Begin signature block
# MIIMSwYJKoZIhvcNAQcCoIIMPDCCDDgCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU8y/xVM4ZZGiXjMQslPRONOuH
# RlugggmqMIIElDCCA3ygAwIBAgIOSBtqBybS6D8mAtSCWs0wDQYJKoZIhvcNAQEL
# BQAwTDEgMB4GA1UECxMXR2xvYmFsU2lnbiBSb290IENBIC0gUjMxEzARBgNVBAoT
# Ckdsb2JhbFNpZ24xEzARBgNVBAMTCkdsb2JhbFNpZ24wHhcNMTYwNjE1MDAwMDAw
# WhcNMjQwNjE1MDAwMDAwWjBaMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFs
# U2lnbiBudi1zYTEwMC4GA1UEAxMnR2xvYmFsU2lnbiBDb2RlU2lnbmluZyBDQSAt
# IFNIQTI1NiAtIEczMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAjYVV
# I6kfU6/J7TbCKbVu2PlC9SGLh/BDoS/AP5fjGEfUlk6Iq8Zj6bZJFYXx2Zt7G/3Y
# SsxtToZAF817ukcotdYUQAyG7h5LM/MsVe4hjNq2wf6wTjquUZ+lFOMQ5pPK+vld
# sZCH7/g1LfyiXCbuexWLH9nDoZc1QbMw/XITrZGXOs5ynQYKdTwfmOPLGC+MnwhK
# kQrZ2TXZg5J2Yl7fg67k1gFOzPM8cGFYNx8U42qgr2v02dJsLBkwXaBvUt/RnMng
# Ddl1EWWW2UO0p5A5rkccVMuxlW4l3o7xEhzw127nFE2zGmXWhEpX7gSvYjjFEJtD
# jlK4PrauniyX/4507wIDAQABo4IBZDCCAWAwDgYDVR0PAQH/BAQDAgEGMB0GA1Ud
# JQQWMBQGCCsGAQUFBwMDBggrBgEFBQcDCTASBgNVHRMBAf8ECDAGAQH/AgEAMB0G
# A1UdDgQWBBQPOueslJF0LZYCc4OtnC5JPxmqVDAfBgNVHSMEGDAWgBSP8Et/qC5F
# JK5NUPpjmove4t0bvDA+BggrBgEFBQcBAQQyMDAwLgYIKwYBBQUHMAGGImh0dHA6
# Ly9vY3NwMi5nbG9iYWxzaWduLmNvbS9yb290cjMwNgYDVR0fBC8wLTAroCmgJ4Yl
# aHR0cDovL2NybC5nbG9iYWxzaWduLmNvbS9yb290LXIzLmNybDBjBgNVHSAEXDBa
# MAsGCSsGAQQBoDIBMjAIBgZngQwBBAEwQQYJKwYBBAGgMgFfMDQwMgYIKwYBBQUH
# AgEWJmh0dHBzOi8vd3d3Lmdsb2JhbHNpZ24uY29tL3JlcG9zaXRvcnkvMA0GCSqG
# SIb3DQEBCwUAA4IBAQAVhCgM7aHDGYLbYydB18xjfda8zzabz9JdTAKLWBoWCHqx
# mJl/2DOKXJ5iCprqkMLFYwQL6IdYBgAHglnDqJQy2eAUTaDVI+DH3brwaeJKRWUt
# TUmQeGYyDrBowLCIsI7tXAb4XBBIPyNzujtThFKAzfCzFcgRCosFeEZZCNS+t/9L
# 9ZxqTJx2ohGFRYzUN+5Q3eEzNKmhHzoL8VZEim+zM9CxjtEMYAfuMsLwJG+/r/uB
# AXZnxKPo4KvcM1Uo42dHPOtqpN+U6fSmwIHRUphRptYCtzzqSu/QumXSN4NTS35n
# fIxA9gccsK8EBtz4bEaIcpzrTp3DsLlUo7lOl8oUMIIFDjCCA/agAwIBAgIMUfr8
# J+jCyr4Ay7YNMA0GCSqGSIb3DQEBCwUAMFoxCzAJBgNVBAYTAkJFMRkwFwYDVQQK
# ExBHbG9iYWxTaWduIG52LXNhMTAwLgYDVQQDEydHbG9iYWxTaWduIENvZGVTaWdu
# aW5nIENBIC0gU0hBMjU2IC0gRzMwHhcNMTYwNzI4MTA1NjE3WhcNMTkwNzI5MTA1
# NjE3WjCBhzELMAkGA1UEBhMCREUxDDAKBgNVBAgTA05SVzESMBAGA1UEBxMJUGFk
# ZXJib3JuMRkwFwYDVQQKExBOZXQgYXQgV29yayBHbWJIMRkwFwYDVQQDExBOZXQg
# YXQgV29yayBHbWJIMSAwHgYJKoZIhvcNAQkBFhFpbmZvQG5ldGF0d29yay5kZTCC
# ASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJWtx+QDzgovn6AmkJ8UCTNr
# xtFJbRCHKNkfev6k35mMkNlibsVnFxooABDKSvaB21nXojMz63g+KLUEN5S4JiX3
# FKq5h2XahwWHvar/r2HMK2uJZ76360ePhuSZTnkifsxvwNxByQ9ot2S1O40AyVU5
# xfEUsBh7vVADMbjqBVlXuNAfsfpfvgjoR0CsOfgKk0CEDZ1wP0bXIkrk021a7lAO
# Yq9kqVDFv8K8O5WYvNcvbtAg3QW5JEaFnM3TMaOOSaWZMmIo7lw3e+B8rqknwmcS
# 66W2E0uayJXKqh/SXfS/xCwO2EzBT9Q1x0XiFR1LlEHQ0T/tfenBUlefIxfDZnEC
# AwEAAaOCAaQwggGgMA4GA1UdDwEB/wQEAwIHgDCBlAYIKwYBBQUHAQEEgYcwgYQw
# SAYIKwYBBQUHMAKGPGh0dHA6Ly9zZWN1cmUuZ2xvYmFsc2lnbi5jb20vY2FjZXJ0
# L2dzY29kZXNpZ25zaGEyZzNvY3NwLmNydDA4BggrBgEFBQcwAYYsaHR0cDovL29j
# c3AyLmdsb2JhbHNpZ24uY29tL2dzY29kZXNpZ25zaGEyZzMwVgYDVR0gBE8wTTBB
# BgkrBgEEAaAyATIwNDAyBggrBgEFBQcCARYmaHR0cHM6Ly93d3cuZ2xvYmFsc2ln
# bi5jb20vcmVwb3NpdG9yeS8wCAYGZ4EMAQQBMAkGA1UdEwQCMAAwPwYDVR0fBDgw
# NjA0oDKgMIYuaHR0cDovL2NybC5nbG9iYWxzaWduLmNvbS9nc2NvZGVzaWduc2hh
# MmczLmNybDATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQUZLedJVdZSZd5
# lwNJFEgIc8KbEFEwHwYDVR0jBBgwFoAUDzrnrJSRdC2WAnODrZwuST8ZqlQwDQYJ
# KoZIhvcNAQELBQADggEBADYcz/+SCP59icPJK5w50yiTcoxnOtoA21GZDpt4GGVf
# RQJDWCDJMkU62xwu5HzqwimbwmBykrAf5Log1fLbggI83zIE4sMjkUe/BnnHpHgK
# LYv+3eLEwglMw/6Gmlq9IqNSD8YmTncGZFoFhrCrgAZUkA6RiVxuZrx2wiluueBI
# vfGs+tRA+7Tgx6Ed9kBybnc+xbAiTCNIcSo9OkPZfc3Q9saMgjIehBMXHLgMdrhv
# N5HXv/r4+aZ6asgv3ggArHrS1Pxp0f60hooVK4bA4Ph1td6YZ5lf8HA4uMmHvOjQ
# iNS0UjXqu5Vs6leIRM3pBjuX45xL6ydUsMlLhZQfansxggILMIICBwIBATBqMFox
# CzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMTAwLgYDVQQD
# EydHbG9iYWxTaWduIENvZGVTaWduaW5nIENBIC0gU0hBMjU2IC0gRzMCDFH6/Cfo
# wsq+AMu2DTAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZ
# BgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYB
# BAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUc/IiQkYxZtf9ZKc6LLLEviow2qowDQYJ
# KoZIhvcNAQEBBQAEggEAMFsEA+AZxcWfzkvTw/zdumc3J5pPcWwwjH9/vDbw3Kv4
# pPp5HEC4MlJEuh0yieLQ+I2v3SWunbiHkPywnBnlzJyNxTNucUDIFx9T5h6T0823
# thAkiId7SaggQTpiL2EUtxtX8mf56I+jMZut4wKWHVT3Rg30gn53PhqzB5lPz1DL
# cH4VPTBTol3m4WrTBmyyiZOWixV+LLd2nxEOC8jfwJOUywV7IbgHqWFfyZIrMk46
# XKUzLOzrbDEH1K9GJgMomOS/h6VdZQdLjW7D2x9bGB2GUcSKl/A0iQEYVuAZF2zZ
# 222CBOp6eN693QVNaqWNzrOtATwT/tslrcQrUZSYog==
# SIG # End signature block
