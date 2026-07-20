param (
	[Parameter(Mandatory=$true, ParameterSetName="default")][string] $ExportFilePath,
	# only needed for v14 with enabled provider mode
	[Parameter(Mandatory=$false)][string] $TenantPrimaryDomain,
	# Optional credential. When supplied, BASIC authentication is used. Omit it to use integrated (NTLM) auth, or force NTLM with -UseIntegratedAuth.
	[Parameter(Mandatory=$false)][PSCredential] $Credential,
	[Parameter(Mandatory=$false)][string] $CredentialFileLocation = "$PSScriptRoot\NspBasicAuthCredential.bin",
	[Parameter(Mandatory=$true, ParameterSetName="setCredentialSet")][switch] $SetCredential,
	[Parameter(Mandatory=$false)][switch] $UseIntegratedAuth
)

if (-not [System.IO.Path]::IsPathRooted($CredentialFileLocation)) {
	$CredentialFileLocation = Join-Path -Path $PSScriptRoot -ChildPath $CredentialFileLocation
}

function Set-loginPass {
	Add-Type -AssemblyName System.Security
	$username = Read-Host -Prompt 'Input your username'
	$securePass = Read-Host -Prompt 'Input your user password' -AsSecureString
	$plainPass = [System.Net.NetworkCredential]::new('', $securePass).Password
	$protected = [System.Security.Cryptography.ProtectedData]::Protect([System.Text.Encoding]::Unicode.GetBytes($plainPass), $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
	$payload = @{ FormatVersion = 1; Username = $username; Password = [System.Convert]::ToBase64String($protected) } | ConvertTo-Json -Compress
	[System.IO.File]::WriteAllText($CredentialFileLocation, $payload, [System.Text.Encoding]::UTF8)
}

function Get-loginPass {
	Add-Type -AssemblyName System.Security
	if (-not (Test-Path $CredentialFileLocation)) { return $null }
	try {
		$payload = [System.IO.File]::ReadAllText($CredentialFileLocation, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
				if ([string]::IsNullOrWhiteSpace($payload.Username) -or [string]::IsNullOrWhiteSpace($payload.Password)) { Write-Warning "Credential file '$CredentialFileLocation' does not use the expected username/password JSON format. Recreate it with -SetCredential. Falling back to integrated authentication."; return $null }
		$rawKey = [System.Security.Cryptography.ProtectedData]::Unprotect([System.Convert]::FromBase64String($payload.Password), $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
		$securePassword = ConvertTo-SecureString -String ([System.Text.Encoding]::Unicode.GetString($rawKey)) -AsPlainText -Force
		return New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $payload.Username, $securePassword
	}
	catch { Write-Warning "Credential file '$CredentialFileLocation' cannot be loaded: $($_.Exception.Message) Recreate it with -SetCredential if it was created on another machine or with an older password-only format. Falling back to integrated authentication."; return $null }
}

if($SetCredential){ Set-loginPass; EXIT }
$resolvedCredential = $null
if (-not $UseIntegratedAuth) {
	if ($Credential) { $resolvedCredential = $Credential }
	else { $resolvedCredential = Get-loginPass }
}

$minver = [Version]4.0.0;
if ($PSVersionTable.PSVersion -gt $minver)
{
	$nspVersion = (Get-ItemProperty -Path HKLM:\SOFTWARE\NoSpamProxy\Components -ErrorAction SilentlyContinue).'Intranet Role'
	if ($nspVersion -gt '14.0') {
		try {
						$connectParams = @{ IgnoreServerCertificateErrors = $true; ErrorAction = 'Stop' }; if ($resolvedCredential) { $connectParams['Credential'] = $resolvedCredential.GetNetworkCredential() }
			Connect-Nsp @connectParams
		} catch {
			$e = $_
			Write-Warning "Not possible to connect with the NoSpamProxy. Please check the error message below."
			$e |Format-List * -Force
			EXIT
		}
		if ($(Get-NspIsProviderModeEnabled) -eq $true) {
			if ($null -eq $TenantPrimaryDomain -OR $TenantPrimaryDomain -eq "") {
				Write-Host "Please provide a TenantPrimaryDomain to run this script with NoSpamProxy v14 in provider mode."
				EXIT
			} else {
				# NSP v14 has a new authentication mechanism, Connect-Nsp is required to authenticate properly
				# -IgnoreServerCertificateErrors allows the usage of self-signed certificates
								$connectParams = @{ IgnoreServerCertificateErrors = $true; PrimaryDomain = $TenantPrimaryDomain }; if ($resolvedCredential) { $connectParams['Credential'] = $resolvedCredential.GetNetworkCredential() }
				Connect-Nsp @connectParams
			}
		}
	}
$features = Get-nspfeatureusage;
$protcounter = 0;
$protarray = @();
$enccounter = 0;
$encarray = @();
$lfcounter = 0;
$lfarray = @();
$disccounter = 0;
$discarray = @();
$sandboxcounter = 0;
$sandboxarray = @();
foreach ($user in $features){
	if ($user.Protection -gt 0)
	{
		$protcounter+=1;
		$protarray+=$user.DisplayName;
	}
	if ($user.Encryption -gt 0)
	{
		$enccounter+=1;
		$encarray+=$user.DisplayName;
	}	
	if ($user.LargeFiles -gt 0)
	{
		$lfcounter+=1;
		$lfarray+=$user.DisplayName;
	}	
	if ($user.Disclaimer -gt 0)
	{
		$disccounter+=1;
		$discarray+=$user.DisplayName;
	}	
	if ($user.FilesUploadedToSandbox -gt 0)
	{
		$sandboxcounter+=$user.FilesUploadedToSandbox;
		$sandboxarray+=$user.DisplayName + " - " + $user.FilesUploadedToSandbox + " Uploaded File(s)";
	}	
}
# the $false prevents UTF8 with BOM
$encoding = New-Object System.Text.UTF8Encoding $false
$stream = New-Object System.IO.StreamWriter $ExportFilePath, $false, $encoding
$stream.WriteLine("Feature-Usage:")
$stream.Write("`r`n")
$stream.WriteLine("Protection: " + $protcounter + " User")
$stream.WriteLine("-----------------------------")
$stream.WriteLine(($protarray | Out-String))
$stream.WriteLine("`r`n`r`n")
$stream.WriteLine("Encryption: " + $enccounter + " User")
$stream.WriteLine("-----------------------------")
$stream.WriteLine(($encarray | Out-String))
$stream.WriteLine("`r`n`r`n")
$stream.WriteLine("LargeFiles: " + $lfcounter + " User")
$stream.WriteLine("-----------------------------")
$stream.WriteLine(($lfarray | Out-String))
$stream.WriteLine("`r`n`r`n")
$stream.WriteLine("Disclaimer: " + $disccounter + " User")
$stream.WriteLine("-----------------------------")
$stream.WriteLine(($discarray | Out-String))
$stream.WriteLine("`r`n`r`n")
$stream.WriteLine("Sandbox: " + $sandboxcounter + " Uploaded Files")
$stream.WriteLine("-----------------------------")
$stream.WriteLine(($sandboxarray | Out-String))
$stream.Close()
}
else 
{
Write-Host "Eror: PowerShell version is to old. Please Update PowrShell to version 4.0.0"
}