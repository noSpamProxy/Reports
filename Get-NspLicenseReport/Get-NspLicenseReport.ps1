param (
	[Parameter(Mandatory=$true)][string] $ExportFilePath
)
$minver = [Version]4.0.0;
if ($PSVersionTable.PSVersion -gt $minver)
{
	$nspVersion = (Get-ItemProperty -Path HKLM:\SOFTWARE\NoSpamProxy\Components -ErrorAction SilentlyContinue).'Intranet Role'
	if ($nspVersion -gt '14.0') {
		# NSP v14 has a new authentication mechanism, Connect-Nsp is required to authenticate properly
		# -IgnoreServerCertificateErrors allows the usage of self-signed certificates
		Connect-Nsp -IgnoreServerCertificateErrors
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