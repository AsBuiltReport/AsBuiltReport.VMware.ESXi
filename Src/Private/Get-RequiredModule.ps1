function Get-RequiredModule {
    <#
    .SYNOPSIS
    Function to check if the required PowerShell module is installed
    .DESCRIPTION
    Function to check if the required PowerShell module is installed
    .NOTES
        Version:        0.2.0
        Author:         Tim Carman
        Twitter:        @tpcarman
        Github:         tpcarman
    .PARAMETER Name
    The name of the required PowerShell module
    .PARAMETER Version
    The version of the required PowerShell module
    #>
    [CmdletBinding()]
    Param (
        [CmdletBinding()]
        [Parameter(Mandatory = $true, ValueFromPipeline = $false)]
        [ValidateNotNullOrEmpty()]
        [String]$Name,

        [CmdletBinding()]
        [Parameter(Mandatory = $true, ValueFromPipeline = $false)]
        [ValidateNotNullOrEmpty()]
        [String]$Version
    )

    # Convert required version to a [Version] object
    $RequiredVersion = [Version]$Version

    # Find the latest installed version of the module
    $InstalledModule = Get-Module -ListAvailable -Name $Name |
        Sort-Object -Property Version -Descending |
        Select-Object -First 1

    if ($null -eq $InstalledModule) {
        throw "$Name $Version or higher is required. Run 'Install-Module -Name $Name -MinimumVersion $Version -Force' to install the required modules."
    }

    # Convert installed version to a [Version] object
    $InstalledVersion = [Version]$InstalledModule.Version

    Write-PScriboMessage -Plugin "Module" -IsWarning "$($InstalledModule.Name) $InstalledVersion is currently installed."

    if ($InstalledVersion -lt $RequiredVersion) {
        throw "$Name $Version or higher is required. Run 'Update-Module -Name $Name -MinimumVersion $Version -Force' to update the required modules."
    }
}