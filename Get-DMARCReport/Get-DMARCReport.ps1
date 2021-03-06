Param
    (
	[Parameter(Mandatory=$true)][string] $SMTPHost,
	[Parameter(Mandatory=$false)][int] $NumberOfDaysToReport = 7,
	[Parameter(Mandatory=$false)][string] $ReportSender = "NoSpamProxy Report Sender <nospamproxy@example.com>",
	[Parameter(Mandatory=$false)][string] $ReportSubject = "DMARC Auswertung der letzten Woche",
	[Parameter(Mandatory=$true)][string] $ReportRecipient
	)

$dateStart = (Get-Date).AddDays(-$NumberOfDaysToReport)
$reportFileName = $Env:TEMP + "\dmarc-report.html"
$XMLreportFileName = $Env:TEMP + "\dmarc-report.xml"
$CSVreportFilename = $Env:TEMP + "\dmarc-report.csv"
$senderDomain = $null
$dkimsenderDomain = $null
$messageID = $null
$spfStatus = $null
$dkimStatus = $null
$dmarcStatus = $null
$wouldhavepassed = $null
$htmlbody2 = $null
$htmlout = $null
$elements = @()

Write-Host "Getting MessageTrackInformation.."
$messageTracks = Get-NSPMessageTrack -Status Success -From $dateStart -Directions FromExternal -WithOperations -WithActions -WithAddresses
Write-Host "Done."
Write-Host "Creating Report...."

foreach ($messageTrack in $messageTracks){
	$operations = $messageTrack.Operations
	foreach ($operation in $operations ){
		if ($operation.Operation.Type -eq "DmarcValidationEntry"){
			$dmarcCheckResult = $operation.Operation.Data
			$dmarchCheckDetails = $dmarcCheckResult |ConvertFrom-Json
			$addresses = $messageTrack.Addresses
			$senderDomain = ($addresses | ? {$_.AddressType -eq "Sender"} | select "Address").Address
			$messageID = $messageTrack.MessageID
			$emailcreated = $messageTrack.Sent.LocalDateTime
			$dmarceffectivePolicy = $dmarchCheckDetails.EffectivePolicy
			$dkimsenderDomain = $dmarchCheckDetails.OrganizationalDomain
			$dmarcsenderDomain = $dmarchCheckDetails.Rfc5322FromDomain
			$dmarcvalidationResult = $dmarchCheckDetails.ValidationResult
			$dkimResult = $dmarchCheckDetails.DkimResult
			$spfResult = $dmarchCheckDetails.SpfResult
			$dkimAlignment = $dmarchCheckDetails.DkimAlignment
			$spfAlignment = $dmarchCheckDetails.SpfAlignment
			$applicablePolicy = $dmarchCheckDetails.ApplicablePolicy
			$dmarcisvalid = $dmarchCheckDetails.validationResult
			foreach ($spfCheckEntry in $spfResult){
				if ($spfCheckEntry.Result -eq "Pass")
				{
				$spfResultDetail +=($spfCheckEntry.Domain +": OK")
				break
				}
				elseif ($spfCheckEntry.Result -ne "None")
				{
				$spfResultDetail +=($spfCheckEntry.Domain +": Failure<br>")
				}
				else
				{
				$spfResultDetail +=($spfCheckEntry.Domain +": None<br>")
				}
			}
			if ($dkimResult.Failures -eq "None"){
				$dkimResultDetail = "DKIM Signature present and OK"
			}
			else
			{
				$dkimResultDetail = "DKIM Signature not present or wrong"
			}
			$htmlbody2 +=("<tr><td><h3> "+$senderDomain +"</h3></td><td><h3> " +$dmarcsenderDomain +"</h3></td><td><h3>" +$spfAlignment +"</h3></td><td><h3>" +$spfResultDetail +"</h3></td><td><h3>"+$dkimAlignment +"</h3></td><td><h3>"+$dkimResultDetail +"</h3></td><td><h3>"+$dmarcvalidationResult +"</h3></td><td><h3>"+$applicablePolicy +"</h3></td><td><h3>"+$dmarceffectivePolicy +"</h3></td><td><h3> "+$emailcreated +"</h3></td><td><h3>"+$messageID +"</h3></td></tr>")
			$o = New-Object PsObject
			$o | Add-Member NoteProperty SenderDomain $senderDomain
			$o | Add-Member NoteProperty DMARCSenderDomain $dmarcsenderDomain
			$o | Add-Member NoteProperty SPFAlignment $spfAlignment
			$o | Add-Member NoteProperty SPFCheckResults $spfResultDetail
			$o | Add-Member NoteProperty DKIMAlignment $dkimAlignment
			$o | Add-Member NoteProperty DKIMCheckResults $dkimResultDetail
			$o | Add-Member NoteProperty DMARCValidationResult $dmarcvalidationResult
			$o | Add-Member NoteProperty ApplicablePolicy $applicablePolicy
			$o | Add-Member NoteProperty WhatIf $dmarceffectivePolicy
			
			$elements +=$o
			$spfResultDetail = $null
		}
	}
}

Write-Host "Done."
Write-Host "Generating and sending report"

$htmlbody1 ="<html>
			<head>
				<title>DMARC Auswertung der letzten Woche</title>
				<style>
	      			table, td, th { border: 1px solid black; border-collapse: collapse; white-space:nowrap; }
					#headerzeile         {background-color: #DDDDDD;}
	    		</style>
			</head>
		<body style=font-family:arial>
			<h1>DMARC Auswertung der letzten Woche</h1>
			<br>
			<table>
				<tr id=headerzeile>
					<td><h3>Sender-Domain</h3></td><td><h3>DMARC-Sender-Domain</h3></td><td><h3>SPF-Alignment</h3></td><td><h3>SPF-Check-Detail</h3></td><td><h3>DKIM-Alignment</h3></td><td><h3>DKIM-Check-Detail</h3></td><td><h3>DMARC Check</h3></td><td><h3>DMARC Policy</h3></td><td><h3>What if</h3></td><td><h3>Datum</h3></td><td><h3>Message ID</h3></td>
				</tr>
				"

$htmlbody3="</table>
		</body>
		</html>"

$htmlout=$htmlbody1+$htmlbody2+$htmlbody3
$htmlout | Out-File $reportFileName
(ConvertTo-Xml -notypeinformation $elements).OuterXml |out-file $XMLreportFileName
$elements.GetEnumerator() |Export-Csv -Path $CSVreportFilename -NoTypeInformation

Send-MailMessage -SmtpServer $SmtpHost -From $ReportSender -To $ReportRecipient -Subject $ReportSubject -Body "Im Anhang dieser E-Mail finden Sie den Bericht mit der DMARC-Auswertung der letzten Woche." -Attachments ($reportFileName, $XMLreportFileName, $CSVreportFilename)

Write-Host "Done."
Write-Host "Doing some cleanup...."
Remove-Item $reportFileName
Remove-Item $XMLreportFileName
Remove-Item $CSVreportFilename
Write-Host "Done."

# SIG # Begin signature block
# MIIMSwYJKoZIhvcNAQcCoIIMPDCCDDgCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU/thXQ7AI1EKEc0TB6gHz/uip
# Z8OgggmqMIIElDCCA3ygAwIBAgIOSBtqBybS6D8mAtSCWs0wDQYJKoZIhvcNAQEL
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
# BAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUjh+rVLtxuCbFbzgsCeFHyUcpfEEwDQYJ
# KoZIhvcNAQEBBQAEggEAAdwwMsAwrViNmCawRvWLXARyY6D/OWSOkw6use3INKVg
# 7VkTwdaaHxcJ2Rkw8mNOgAWkUdDm7uKE0zfsBoY1zOoHXzumoWY7oPqgXYTRkQkA
# 8rW4Ofa+Ve078+hyM6LD7UbwY7taSgnerfH90Pc7qXeQjIv7ZZ9ZKK/YlZeJmgSD
# zgyG7rJ7b9UgUZnGHvqyAK+ewPp6WY8HjtVSqowinsdVvFDmcd3A+HFetFbeZnnz
# UskQSWRMLf8BSqDoaMpof7bDBUlUxNvJye3hCIB2/Rf9fPJG6Sf8whmVobwfsH+C
# kQ3HDEFrbr5azZ1haAjZpbWrkHWBkypus9WsYCLkOg==
# SIG # End signature block
