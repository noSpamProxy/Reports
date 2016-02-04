param (
	[Parameter(Mandatory=$false)][int] $NumberOfDaysToReport = 1,
	[Parameter(Mandatory=$true)][string] $SmtpHost,
	[Parameter(Mandatory=$true)][string] $ReportSender,
	[Parameter(Mandatory=$true)][string] $ReportRecipientt,
	[Parameter(Mandatory=$false)][string] $ReportSubject = "Auswertung"
)

$reportFileName = $Env:TEMP + "\reject-analysis.html"
$totalRejected = 0
$tempRejected = 0
$permanentRejected = 0
$rdnsTempRejected = 0
$rblRejected = 0
$cyrenSpamRejected = 0
$cyrenAVRejected = 0
$surblRejected = 0
$characterSetRejected = 0
$headerFromRejected = 0
$wordRejected = 0
$rdnsPermanentRejected = 0
$decryptPolicyRejected = 0
$onBodyRejected = 0
$onEnvelopeRejected = 0
$dateStart = (Get-Date).AddDays(-$NumberOfDaysToReport)

$messageTracks = Get-NSPMessageTrack -From $dateStart -Status TemporarilyBlocked -Directions FromExternal

foreach ($item in $messageTracks)
{
	$totalRejected++
	$tempRejected++
	$tempvalidationentries = $item.Details.ValidationResult.ValidationEntries
	foreach ($tempvalidationentry in $tempvalidationentries)
	{
		if (($tempvalidationentry.Id -eq "reverseDnsLookup") -and ($tempvalidationentry.Decision -eq "RejectTemporarily" ))
		{
			$rdnsTempRejected++
			$onEnvelopeRejected++
		}
	}
}

$messageTracks = Get-NSPMessageTrack -From $dateStart -Status PermanentlyBlocked -Directions FromExternal

foreach ($item in $messageTracks)
{
	$totalRejected++
	$permanentRejected++
	$permanentvalidationentries = $item.Details.ValidationResult.ValidationEntries
	foreach ($permanentvalidationentry in $permanentvalidationentries)
	{
		if ($permanentvalidationentry.Id -eq "realtimeBlocklist" -and $permanentvalidationentry.Scl -gt 0)
		{
			$rblRejected++
			$onEnvelopeRejected++
		}
		if ($permanentvalidationentry.Id -eq "cyrenAction" -and $permanentvalidationentry.Decision -notcontains "Pass")
		{
			$cyrenAVRejected++
			$onBodyRejected++
		}
		if ($permanentvalidationentry.Id -eq "surblFilter" -and $permanentvalidationentry.Scl -gt 0)
		{
			$surblRejected++
			$onBodyRejected++
		}
		if ($permanentvalidationentry.Id -eq "cyrenFilter" -and $permanentvalidationentry.Scl -gt 0)
		{
			$cyrenSpamRejected++
			$onBodyRejected++
		}
		if ($permanentvalidationentry.Id -eq "characterSetFilter" -and $permanentvalidationentry.Scl -gt 0)
		{
			$characterSetRejected++
			$onBodyRejected++
		}
		if ($permanentvalidationentry.Id -eq "ensureHeaderFromIsExternal" -and $permanentvalidationentry.Scl -gt 0)
		{
			$headerFromRejected++
			$onBodyRejected++
		}
		if ($permanentvalidationentry.Id -eq "wordFilter" -and $permanentvalidationentry.Scl -gt 0)
		{
			$wordRejected++
			$onBodyRejected++
		}
		if (($permanentvalidationentry.Id -eq "reverseDnsLookup") -and ($permanentvalidationentry.Decision -eq "RejectPermanent" ))
		{
			$rdnsPermanentRejected++
			$onEnvelopeRejected++
		}
		if (($permanentvalidationentry.Id -eq "validateSignatureAndDecrypt") -and ($permanentvalidationentry.Decision -notcontains "Pass" ))
		{
			$decryptPolicyRejected++
			$onBodyRejected++
		}
	}
}

Write-Host "------------------------------"
Write-Host "OnEnvelope Reject:" $onEnvelopeRejected
Write-Host ""
Write-Host "RDNS PermanentReject:" $rdnsPermanentRejected
Write-Host "RDNS TempReject:" $rdnsTempRejected
Write-Host "RBL rejected:" $rblRejected
Write-Host "------------------------------"
Write-Host "OnBody Reject:" $onBodyRejected
Write-Host ""
Write-Host "CyrenSpam Reject:" $cyrenSpamRejected
Write-Host "CyrenAV Reject:" $cyrenAVRejected
Write-Host "SURBL Reject:" $surblRejected
Write-Host "CharacterSet Reject:" $characterSetRejected
Write-Host "HeaderFrom Reject:" $headerFromRejected
Write-Host "Word Reject:" $wordRejected
Write-Host "DecryptPolicy Reject:" $decryptPolicyRejected
Write-Host "------------------------------"
Write-Host "TemporaryReject Total:" $tempRejected
Write-Host "PermanentReject Total:" $permanentRejected
Write-Host "TotalReject:" $totalRejected

$htmlout = "<html>
		<head>
			<title>Auswertung der abgewiesenen E-Mails</title>
			<style>
      			table, td, th { border: 1px solid black; border-collapse: collapse; }
				#headerzeile         {background-color: #DDDDDD;}
    		</style>
		</head>
	<body style=font-family:arial>
		<h1>Auswertung der abgewiesenen E-Mails</h1>
		<br>
		<table>
			<tr id=headerzeile><td colspan=2><b>On Envelope Level: " +$onEnvelopeRejected +"</b></td></tr>
			<tr id=headerzeile><td><b>Filter</b></td><td width=250px><b>Anzahl</b></td></td></tr>
			<tr><td>RDNS PermanentReject</td><td>" + $rdnsPermanentRejected +"</td></tr>
			<tr><td>RDNS TempReject</td><td>" + $rdnsTempRejected +"</td></tr>
			<tr><td>Realtime Blocklists</td><td>" + $rblRejected +"</td></tr>
			<tr><td colspan=2>&nbsp;</td></tr>
			<tr id=headerzeile><td colspan=2><b>On Body Level: " +$onBodyRejected +"<b></td></tr>
			<tr id=headerzeile><td><b>Filter</b></td><td width=250px><b>Anzahl</b></td></td></tr>
			<tr><td>Cyren AntiSpam</td><td>" + $cyrenSpamRejected +"</td></tr>
			<tr><td>Cyren Premium AntiVirus</td><td>" + $cyrenAVRejected +"</td></tr>
			<tr><td>Spam URI Realtime Blocklists</td><td>" + $surblRejected +"</td></tr>
			<tr><td>Allowed Unicode Character Sets</td><td>" + $characterSetRejected +"</td></tr>
			<tr><td>Prevent Owned Domains in HeaderFrom</td><td>" + $headerFromRejected +"</td></tr>
			<tr><td>Word Matching</td><td>" + $wordRejected +"</td></tr>
			<tr><td>DecryptPolicy Reject</td><td>" + $decryptPolicyRejected +"</td></tr>
			<tr><td colspan=2>&nbsp;</td></tr>
			<tr><td><h3>TemporaryReject Total</h3></td><td><h3>" + $tempRejected +"</h3></td></tr>
			<tr><td><h3>PermanentReject Total</h3></td><td><h3>" + $permanentRejected +"</h3></td></tr>
			<tr><td><h3>Reject Total</h3></td><td><h3>" + $totalRejected +"</h3></td></tr>
		</table>
	</body>
	</html>"

$htmlout | Out-File $reportFileName
Send-MailMessage -SmtpServer $SmtpHost -From $ReportSender -To $ReportRecipient -Subject $ReportSubject -Body "Im Anhang dieser E-Mail finden Sie den Bericht mit der Auswertung." -Attachments $reportFileName
Remove-Item $reportFileName

# SIG # Begin signature block
# MIIMNAYJKoZIhvcNAQcCoIIMJTCCDCECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUiz2NnT8gOLHUXyrOFHQxacJ3
# agCgggmdMIIEmTCCA4GgAwIBAgIQcaC3NpXdsa/COyuaGO5UyzANBgkqhkiG9w0B
# AQsFADCBqTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDHRoYXd0ZSwgSW5jLjEoMCYG
# A1UECxMfQ2VydGlmaWNhdGlvbiBTZXJ2aWNlcyBEaXZpc2lvbjE4MDYGA1UECxMv
# KGMpIDIwMDYgdGhhd3RlLCBJbmMuIC0gRm9yIGF1dGhvcml6ZWQgdXNlIG9ubHkx
# HzAdBgNVBAMTFnRoYXd0ZSBQcmltYXJ5IFJvb3QgQ0EwHhcNMTMxMjEwMDAwMDAw
# WhcNMjMxMjA5MjM1OTU5WjBMMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMdGhhd3Rl
# LCBJbmMuMSYwJAYDVQQDEx10aGF3dGUgU0hBMjU2IENvZGUgU2lnbmluZyBDQTCC
# ASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJtVAkwXBenQZsP8KK3TwP7v
# 4Ol+1B72qhuRRv31Fu2YB1P6uocbfZ4fASerudJnyrcQJVP0476bkLjtI1xC72Ql
# WOWIIhq+9ceu9b6KsRERkxoiqXRpwXS2aIengzD5ZPGx4zg+9NbB/BL+c1cXNVeK
# 3VCNA/hmzcp2gxPI1w5xHeRjyboX+NG55IjSLCjIISANQbcL4i/CgOaIe1Nsw0Rj
# gX9oR4wrKs9b9IxJYbpphf1rAHgFJmkTMIA4TvFaVcnFUNaqOIlHQ1z+TXOlScWT
# af53lpqv84wOV7oz2Q7GQtMDd8S7Oa2R+fP3llw6ZKbtJ1fB6EDzU/K+KTT+X/kC
# AwEAAaOCARcwggETMC8GCCsGAQUFBwEBBCMwITAfBggrBgEFBQcwAYYTaHR0cDov
# L3QyLnN5bWNiLmNvbTASBgNVHRMBAf8ECDAGAQH/AgEAMDIGA1UdHwQrMCkwJ6Al
# oCOGIWh0dHA6Ly90MS5zeW1jYi5jb20vVGhhd3RlUENBLmNybDAdBgNVHSUEFjAU
# BggrBgEFBQcDAgYIKwYBBQUHAwMwDgYDVR0PAQH/BAQDAgEGMCkGA1UdEQQiMCCk
# HjAcMRowGAYDVQQDExFTeW1hbnRlY1BLSS0xLTU2ODAdBgNVHQ4EFgQUV4abVLi+
# pimK5PbC4hMYiYXN3LcwHwYDVR0jBBgwFoAUe1tFz6/Oy3r9MZIaarbzRutXSFAw
# DQYJKoZIhvcNAQELBQADggEBACQ79degNhPHQ/7wCYdo0ZgxbhLkPx4flntrTB6H
# novFbKOxDHtQktWBnLGPLCm37vmRBbmOQfEs9tBZLZjgueqAAUdAlbg9nQO9ebs1
# tq2cTCf2Z0UQycW8h05Ve9KHu93cMO/G1GzMmTVtHOBg081ojylZS4mWCEbJjvx1
# T8XcCcxOJ4tEzQe8rATgtTOlh5/03XMMkeoSgW/jdfAetZNsRBfVPpfJvQcsVncf
# hd1G6L/eLIGUo/flt6fBN591ylV3TV42KcqF2EVBcld1wHlb+jQQBm1kIEK3Osgf
# HUZkAl/GR77wxDooVNr2Hk+aohlDpG9J+PxeQiAohItHIG4wggT8MIID5KADAgEC
# AhAh36cYPt9rQMtVY5K+Zf5LMA0GCSqGSIb3DQEBCwUAMEwxCzAJBgNVBAYTAlVT
# MRUwEwYDVQQKEwx0aGF3dGUsIEluYy4xJjAkBgNVBAMTHXRoYXd0ZSBTSEEyNTYg
# Q29kZSBTaWduaW5nIENBMB4XDTE1MDkyMTAwMDAwMFoXDTE2MDkyMDIzNTk1OVow
# gZUxCzAJBgNVBAYTAkRFMRwwGgYDVQQIExNOb3JkcmhlaW4gV2VzdGZhbGVuMRIw
# EAYDVQQHFAlQYWRlcmJvcm4xKTAnBgNVBAoUIE5ldCBhdCBXb3JrIE5ldHp3ZXJr
# c3lzdGVtZSBHbWJIMSkwJwYDVQQDFCBOZXQgYXQgV29yayBOZXR6d2Vya3N5c3Rl
# bWUgR21iSDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAMIsja4vgLIG
# rpdvUkdCsS8HCjLwaFXt8TIXG8NYIed1aaG+tV0cmScVlsVRUSRfdKVlaTrg7ZDa
# v17t5rFle0fI8XlaMTt86mp8ujdo+svKpHSXiWL51LiADwRETzqIQfUXkdZqgXGg
# wBrTu0zzIH6NvRm7o7o43sSw5rHTHyKPJUDNEE+gAfPsH/69xDmMuH/2r6iMe5GZ
# dRyAmEtB+sEOdhCIX45gXCEGtc3lPeUDCi4I0P6+oqwHzmgfh3IIBF/PCda4V8yP
# lk65x3+6X1eNox3hWQxNQX2cOx1Yd8yaH9ZYdY8y+RwYauaiGOhzf5XvQtfuka6P
# GR270YqN7/ECAwEAAaOCAY4wggGKMAkGA1UdEwQCMAAwHwYDVR0jBBgwFoAUV4ab
# VLi+pimK5PbC4hMYiYXN3LcwHQYDVR0OBBYEFH3SkQBtD02UoOt/MNcKFgvst2+/
# MCsGA1UdHwQkMCIwIKAeoByGGmh0dHA6Ly90bC5zeW1jYi5jb20vdGwuY3JsMA4G
# A1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzBzBgNVHSAEbDBqMGgG
# C2CGSAGG+EUBBzACMFkwJgYIKwYBBQUHAgEWGmh0dHBzOi8vd3d3LnRoYXd0ZS5j
# b20vY3BzMC8GCCsGAQUFBwICMCMMIWh0dHBzOi8vd3d3LnRoYXd0ZS5jb20vcmVw
# b3NpdG9yeTAdBgNVHQQEFjAUMA4wDAYKKwYBBAGCNwIBFgMCB4AwVwYIKwYBBQUH
# AQEESzBJMB8GCCsGAQUFBzABhhNodHRwOi8vdGwuc3ltY2QuY29tMCYGCCsGAQUF
# BzAChhpodHRwOi8vdGwuc3ltY2IuY29tL3RsLmNydDANBgkqhkiG9w0BAQsFAAOC
# AQEAjQCSIdnnJcXUpByMElfYuBh0o66Z9D0teIP7tstExgFpUEdV2i1QgftYTod9
# kflbJWL+kreYq0v3Ibi70X2+o46cbKMncZpkuPNgUN91mn5V0B3DONgrE7FYZ2Ts
# JP5PR+wOunVtIaKn3SbOqTocbDx3SLBaGly+bPnh5FqsudhRWqiMKzQHxy3Lh03c
# PYYRkGUjjZekS6s3cYFZremd8TZyZgiU6ifCI8e3wNK1GFv8M7DFYHa0ta27jofc
# DtJW6f0U+8GY99R3HP3B99Lw96Gf3RMjH4ItbpT0vImZLPoA5FyigphBdYnAiZ9N
# Pd0LwA/vo00NG6ZHXUliXjH4UjGCAgEwggH9AgEBMGAwTDELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDHRoYXd0ZSwgSW5jLjEmMCQGA1UEAxMddGhhd3RlIFNIQTI1NiBD
# b2RlIFNpZ25pbmcgQ0ECECHfpxg+32tAy1Vjkr5l/kswCQYFKw4DAhoFAKB4MBgG
# CisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcC
# AQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARYwIwYJKoZIhvcNAQkEMRYE
# FPRdIOHHmvPd50ZtB2hPHoxVJdBjMA0GCSqGSIb3DQEBAQUABIIBAEfgFwH8JkzV
# ngWDnSm/q3AMZUz/Dnyu2VRZLnCZMUclaRczqCOuKMiE7loXvfbTkHlJkhOvfFaC
# CtifZ8MVwbFBMQZtMcrvYVfTTrwIKXknMJu21VZtv9JrUUJCYwT+CMNnY/Xh0JOZ
# qja5JA3fKF2LZF7lS9mK0xI2NG3QBdMrJAQGa32qqsZozH2439k2hymKTkweHGP1
# ihc9cQWnyd1lW3uy7E0ll2f+euGTJViSGW/b2Y8K5LDUEFWMpHrZZaH/la40e97m
# 7QKahRGmD0hthueBtC8k+LvB1f9XzelwAnMuFG69s+mfJJ6FpK9nHlQiYAXk++/Z
# 0cf7ELcZWPs=
# SIG # End signature block
