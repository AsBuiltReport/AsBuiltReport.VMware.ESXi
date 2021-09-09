function Get-VMHostLicense {
    <#
.SYNOPSIS
Function to retrieve VMware ESXi product licensing information.
.DESCRIPTION
Function to retrieve VMware ESXi product licensing information.
.NOTES
Version:        0.1.0
Author:         Tim Carman
Twitter:        @tpcarman
Github:         tpcarman
.PARAMETER VMHost
A vSphere ESXi Host objects
.INPUTS
System.Management.Automation.PSObject.
.OUTPUTS
System.Management.Automation.PSObject.
.EXAMPLE
PS> Get-VMHostLicense -VMHost ESXi01
#>
    [CmdletBinding()][OutputType('System.Management.Automation.PSObject')]

    Param
    (
        [Parameter(Mandatory = $false, ValueFromPipeline = $false)]
        [ValidateNotNullOrEmpty()]
        [PSObject]$VMHost,
        [Parameter(Mandatory = $false, ValueFromPipeline = $false)]
        [Switch]$Licenses
    ) 

    if ($VMHost) {
        $LicenseObject = @()
        $ServiceInstance = Get-View ServiceInstance -Server $ESXi
        $LicenseManager = Get-View $ServiceInstance.Content.LicenseManager -Server $ESXi
        #$LicenseManagerAssign = Get-View $LicenseManager.LicenseAssignmentManager
    
        #$VMHostId = $VMHost.Extensiondata.Config.Host.Value
        #$VMHostAssignedLicense = $LicenseManagerAssign.QueryAssignedLicenses($VMHostId)    
        $VMHostLicense = $LicenseManager.Licenses
        $VMHostLicenseExpiration = ($VMHostLicense.Properties | Where-Object { $_.Key -eq 'expirationDate' } | Select-Object Value).Value
        if ($VMHostLicense.LicenseKey -and $Options.ShowLicenseKeys) {
            $VMHostLicenseKey = $VMHostLicense.LicenseKey
        } else {
            $VMHostLicenseKey = "*****-*****-*****" + $VMHostLicense.LicenseKey.Substring(17)
        }
        $LicenseObject = [PSCustomObject]@{                               
            Product = $VMHostLicense.Name 
            LicenseKey = $VMHostLicenseKey
            Expiration =
            if ($VMHostLicenseExpiration -eq $null) {
                "Never" 
            } elseif ($VMHostLicenseExpiration -gt (Get-Date)) {
                $VMHostLicenseExpiration.ToShortDateString()
            } else {
                "Expired"
            }
        }
    }        
    Write-Output $LicenseObject
}