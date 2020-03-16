# Get-NspLicenseReport.ps1

Outputs the state of a certificate and its chain to the console.
It is only possible to check certificates which are imported into the NoSpamProxy.
The validation is done for each connected gateway role.

## Usage 

```ps
Get-NspCertificates.ps1 -Thumbprint
```

## Parameters
### Thumbprint
    Set the thumbprint of the certificate which should be checked.
	
## Example

```ps
.\Get-NspCertificates.ps1 -Thumbprint 0F4F9209E172B6D81022C0219CF253EFD29689F6
```
