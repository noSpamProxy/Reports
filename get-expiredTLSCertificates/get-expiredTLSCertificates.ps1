Param
    (
	[Parameter(Mandatory=$true)][string] $SMTPHost,
	[Parameter(Mandatory=$false)][int] $NumberOfDaysToReport = 7,
	[Parameter(Mandatory=$false)][string] $ReportSender = "NoSpamProxy Report Sender <nospamproxy@example.com>",
	[Parameter(Mandatory=$false)][string] $ReportSubject = "Auswertung der abgelaufenen TLS-Zertifikate",
	[Parameter(Mandatory=$true)][string] $ReportRecipient
	)

$dateStart = (Get-Date).AddDays(-$NumberOfDaysToReport)
$actualdate = Get-Date
$reportaddressesFileName = $Env:TEMP + "\expiredcertificates.html"
$recipientDomain = $null
$htmlbody2 = $null
$htmlout = $null
$elements = @()

Write-Host "Getting MessageTrackInformation.."
$messageTracks = Get-NSPMessageTrack -Status Success -From $dateStart -Directions FromLocal |Get-NSPMessageTrackdetails
Write-Host "Done."
Write-Host "Create Reportaddresses-File"

foreach ($messageTrack in $messageTracks)
	{
		$perRecipientStates = $messageTrack.Details.PerRecipientStatus
		foreach ($perRecipientStatus in $perRecipientStates)
			{
			$recipientDomain = $perRecipientStatus.RecipientAddress.Domain
			$connectionSecurityProtocol = $perRecipientStatus.ConnectionSecurity.SslProtocol
			if ($connectionSecurityProtocol -eq 0 ) {continue}
			$remoteIdentity = $perRecipientStatus.ConnectionSecurity.RemoteIdentity
			if ($remoteIdentity -eq $null) {continue}
			$TLSCertificateExpirationDate = $perRecipientStatus.ConnectionSecurity.RemoteIdentity.NotAfter
			$recipientIPAddress = $perRecipientStatus.TargetHost.Address.IPAddressToString
			$certificateName = $perRecipientStatus.ConnectionSecurity.RemoteIdentity.Subject
			if ($TLSCertificateExpirationDate -le $actualdate)
				{
					if($elements | where {$_.recipientDomain -eq $recipientDomain}) {continue}
					
					
					$o = New-Object PsObject
					$o | Add-Member NoteProperty recipientDomain $recipientDomain
					$o | Add-Member NoteProperty expirationDate $TLSCertificateExpirationDate
					$o | Add-Member NoteProperty IPAddress $recipientIPAddress
					$o | Add-Member NoteProperty CertificateSubject $certificateName
					
					$elements +=$o
					
				}
			}
	}

Write-Host "Done."
Write-Host "Generating and sending report"

$htmlbody1 ="<html>
			<head>
				<title>Auswertung der letzten Woche</title>
				<style>
	      			table, td, th { border: 1px solid black; border-collapse: collapse; white-space:nowrap; }
					#headerzeile         {background-color: #DDDDDD;}
	    		</style>
			</head>
		<body style=font-family:arial>
			<h1>Auflisting der E-Mail-Server mit abgelaufenen TLS-Zertifikaten der letzten "+$NumberOfDaysToReport+" Tage</h1>
			<br>
			<table>
				<tr id=headerzeile>
					<td><h3>Empf&auml;nger-Domain</h3></td><td><h3>Ablaufdatum</h3></td><td><h3>IP-Adresse</h3></td><td><h3>Zertifikatsinhaber</h3></td>
				</tr>
				"
$elements | sort recipientDomain | %{
	$htmlbody2 +=("<tr><td>"+$_.recipientDomain +"</td><td>" +$_.expirationDate +"</td><td>" +$_.IPAddress +"</td><td> " +$_.CertificateSubject +"</td></tr>")
}


$htmlbody3="</table>
		</body>
		</html>"

$htmlout=$htmlbody1+$htmlbody2+$htmlbody3
$htmlout | Out-File $reportaddressesFileName

Send-MailMessage -SmtpServer $SmtpHost -From $ReportSender -To $ReportRecipient -Subject $ReportSubject -Body "Im Anhang dieser E-Mail finden Sie die Auflisting der E-Mail-Server mit abgelaufenen TLS-Zertifikaten." -Attachments $reportaddressesFileName

Write-Host "Done."
Write-Host "Doing some cleanup...."
Remove-Item $reportaddressesFileName
Write-Host "Done."
# SIG # Begin signature block
# MIIMNAYJKoZIhvcNAQcCoIIMJTCCDCECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUJRpcm18XvLDDMbpmVLQRPUff
# y2egggmdMIIEmTCCA4GgAwIBAgIQcaC3NpXdsa/COyuaGO5UyzANBgkqhkiG9w0B
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
# FNDIif+PZx02IGGSF7Lusein+eF/MA0GCSqGSIb3DQEBAQUABIIBAGQOD9y9uNjo
# jz0oJnM/A334djGgaIB5akzm0kApLCoAU4gVwQ96iR5thDA0x4qjSmx+Z6SUtLv7
# PnmRFO5DJMbhRvMk2Ehef0h9Skh6Cxx3AeVPOH6aWz9c//tYS5ZQeabfSEMOfB6H
# RVXV9Ho2x30Jq1y743qm8hXYmeiduwK6xwILdmvNbt7IsonJsTgjxW93b80BgKf7
# OlJkzYdr/zPoGu8MaT7u1sodxtEEO3wyP/HI6mm60HEQlGmhF/+ZBq6jrvR4Nzp9
# 8TzHfiRQ4ktyjmJnfXrSWtO/H1IZR5ZtFWP06d2B9AY1tQk2TwtaN8zmJxH7DeJq
# 0+QjZtgHwwo=
# SIG # End signature block
