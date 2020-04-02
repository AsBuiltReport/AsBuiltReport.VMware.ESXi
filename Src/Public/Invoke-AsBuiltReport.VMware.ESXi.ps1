function Invoke-AsBuiltReport.VMware.ESXi {
    <#
    .SYNOPSIS  
        PowerShell script to document the configuration of VMware ESXi servers in Word/HTML/XML/Text formats
    .DESCRIPTION
        Documents the configuration of VMware ESXi servers in Word/HTML/XML/Text formats using PScribo.
    .NOTES
        Version:        1.0.0
        Author:         Tim Carman
        Twitter:        @tpcarman
        Github:         tpcarman
        Credits:        Iain Brighton (@iainbrighton) - PScribo module
    .LINK
        https://github.com/AsBuiltReport/AsBuiltReport.VMware.ESXi
    #>

    param (
        [String[]] $Target,
        [PSCredential] $Credential,
        [String] $StylePath
    )

    # Import JSON Configuration for Options and InfoLevel
    $InfoLevel = $ReportConfig.InfoLevel
    $Options = $ReportConfig.Options

    $TextInfo = (Get-Culture).TextInfo

    # If custom style not set, use default style
    if (!$StylePath) {
        & "$PSScriptRoot\..\..\AsBuiltReport.VMware.ESXi.Style.ps1"
    }

    #region Script Functions
    #---------------------------------------------------------------------------------------------#
    #                                    SCRIPT FUNCTIONS                                         #
    #---------------------------------------------------------------------------------------------#

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
            [PSObject]$vCenter, 
            [PSObject]$VMHost,
            [Parameter(Mandatory = $false, ValueFromPipeline = $false)]
            [Switch]$Licenses
        ) 

        if ($VMHost) {
            $LicenseObject = @()
            $ServiceInstance = Get-View ServiceInstance
            $LicenseManager = Get-View $ServiceInstance.Content.LicenseManager
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

    function Get-VMHostNetworkAdapterCDP {
        <#
    .SYNOPSIS
    Function to retrieve the Network Adapter CDP info of a vSphere host.
    .DESCRIPTION
    Function to retrieve the Network Adapter CDP info of a vSphere host.
    .PARAMETER VMHost
    A vSphere ESXi Host object
    .INPUTS
    System.Management.Automation.PSObject.
    .OUTPUTS
    System.Management.Automation.PSObject.
    .EXAMPLE
    PS> Get-VMHostNetworkAdapterCDP -VMHost ESXi01,ESXi02
    .EXAMPLE
    PS> Get-VMHost ESXi01,ESXi02 | Get-VMHostNetworkAdapterCDP
    #>
        [CmdletBinding()][OutputType('System.Management.Automation.PSObject')]

        Param
        (
            [parameter(Mandatory = $true, ValueFromPipeline = $true)]
            [ValidateNotNullOrEmpty()]
            [PSObject[]]$VMHosts   
        )    

        begin {
            $CDPObject = @()
        }

        process {
            try {
                foreach ($VMHost in $VMHosts) {
                    $ConfigManagerView = Get-View $VMHost.ExtensionData.ConfigManager.NetworkSystem
                    $pNics = $ConfigManagerView.NetworkInfo.Pnic
                    foreach ($pNic in $pNics) {
                        $PhysicalNicHintInfo = $ConfigManagerView.QueryNetworkHint($pNic.Device)
                        $Object = [PSCustomObject]@{                            
                            'VMHost' = $VMHost.ExtensionData.Name
                            'Device' = $pNic.Device
                            'Status' = if ($PhysicalNicHintInfo.ConnectedSwitchPort) {
                                'Connected'
                            } else {
                                'Disconnected'
                            }
                            'SwitchId' = $PhysicalNicHintInfo.ConnectedSwitchPort.DevId
                            'Address' = $PhysicalNicHintInfo.ConnectedSwitchPort.Address
                            'VLAN' = $PhysicalNicHintInfo.ConnectedSwitchPort.Vlan
                            'MTU' = $PhysicalNicHintInfo.ConnectedSwitchPort.Mtu
                            'SystemName' = $PhysicalNicHintInfo.ConnectedSwitchPort.SystemName
                            'Location' = $PhysicalNicHintInfo.ConnectedSwitchPort.Location
                            'HardwarePlatform' = $PhysicalNicHintInfo.ConnectedSwitchPort.HardwarePlatform
                            'SoftwareVersion' = $PhysicalNicHintInfo.ConnectedSwitchPort.SoftwareVersion
                            'ManagementAddress' = $PhysicalNicHintInfo.ConnectedSwitchPort.MgmtAddr
                            'PortId' = $PhysicalNicHintInfo.ConnectedSwitchPort.PortId
                        }
                        $CDPObject += $Object
                    }
                }
            } catch [Exception] {
                throw 'Unable to retrieve CDP info'
            }
        }
        end {
            Write-Output $CDPObject
        }
    }

    function Get-InstallDate {
        $esxcli = Get-EsxCli -VMHost $VMHost -V2
        $thisUUID = $esxcli.system.uuid.get.Invoke()
        $decDate = [Convert]::ToInt32($thisUUID.Split("-")[0], 16)
        $installDate = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($decDate))
        [PSCustomObject][Ordered]@{
            Name = $VMHost.ExtensionData.Name
            InstallDate = $installDate
        }
    }

    function Get-Uptime {
        [CmdletBinding()][OutputType('System.Management.Automation.PSObject')]
        Param (
            [Parameter(Mandatory = $false, ValueFromPipeline = $false)]
            [ValidateNotNullOrEmpty()]
            [PSObject]$VMHost, [PSObject]$VM
        )
        $UptimeObject = @()
        $Date = (Get-Date).ToUniversalTime() 
        If ($VMHost) {
            $UptimeObject = Get-View -ViewType hostsystem -Property Name, Runtime.BootTime -Filter @{
                "Name" = "^$($VMHost.ExtensionData.Name)$"
                "Runtime.ConnectionState" = "connected"
            } | Select-Object Name, @{L = 'UptimeDays'; E = { [math]::round(((($Date) - ($_.Runtime.BootTime)).TotalDays), 2) } }, @{L = 'UptimeHours'; E = { [math]::round(((($Date) - ($_.Runtime.BootTime)).TotalHours), 2) } }, @{L = 'UptimeMinutes'; E = { [math]::round(((($Date) - ($_.Runtime.BootTime)).TotalMinutes), 2) } }
        }

        if ($VM) {
            $UptimeObject = Get-View -ViewType VirtualMachine -Property Name, Runtime.BootTime -Filter @{
                "Name" = "^$($VM.Name)$"
                "Runtime.PowerState" = "poweredOn"
            } | Select-Object Name, @{L = 'UptimeDays'; E = { [math]::round(((($Date) - ($_.Runtime.BootTime)).TotalDays), 2) } }, @{L = 'UptimeHours'; E = { [math]::round(((($Date) - ($_.Runtime.BootTime)).TotalHours), 2) } }, @{L = 'UptimeMinutes'; E = { [math]::round(((($Date) - ($_.Runtime.BootTime)).TotalMinutes), 2) } }
        }
        Write-Output $UptimeObject
    }

    function Get-ESXiBootDevice {
        <#
    .NOTES
    ===========================================================================
        Created by:    William Lam
        Organization:  VMware
        Blog:          www.virtuallyghetto.com
        Twitter:       @lamw
    ===========================================================================
    .DESCRIPTION
        This function identifies how an ESXi host was booted up along with its boot
        device (if applicable). This supports both local installation to Auto Deploy as
        well as Boot from SAN.
    .PARAMETER VMHostname
        The name of an individual ESXi host managed by vCenter Server
    .EXAMPLE
        Get-ESXiBootDevice
    .EXAMPLE
        Get-ESXiBootDevice -VMHost esxi-01
    #>
        param(
            [Parameter(Mandatory = $false)][PSObject]$VMHost
        )

        $results = @()
        $esxcli = Get-EsxCli -V2 -VMHost $vmhost
        $bootDetails = $esxcli.system.boot.device.get.Invoke()

        # Check to see if ESXi booted over the network
        $networkBoot = $false
        if ($bootDetails.BootNIC) {
            $networkBoot = $true
            $bootDevice = $bootDetails.BootNIC
        } elseif ($bootDetails.StatelessBootNIC) {
            $networkBoot = $true
            $bootDevice = $bootDetails.StatelessBootNIC
        }

        # If ESXi booted over network, check to see if deployment
        # is Stateless, Stateless w/Caching or Stateful
        if ($networkBoot) {
            $option = $esxcli.system.settings.advanced.list.CreateArgs()
            $option.option = "/UserVars/ImageCachedSystem"
            try {
                $optionValue = $esxcli.system.settings.advanced.list.Invoke($option)
            } catch {
                $bootType = "Stateless"
            }
            $bootType = $optionValue.StringValue
        }

        # Loop through all storage devices to identify boot device
        $devices = $esxcli.storage.core.device.list.Invoke()
        $foundBootDevice = $false
        foreach ($device in $devices) {
            if ($device.IsBootDevice -eq $true) {
                $foundBootDevice = $true

                if ($device.IsLocal -eq $true -and $networkBoot -and $bootType -ne "Stateful") {
                    $bootType = "Stateless Caching"
                } elseif ($device.IsLocal -eq $true -and $networkBoot -eq $false) {
                    $bootType = "Local"
                } elseif ($device.IsLocal -eq $false -and $networkBoot -eq $false) {
                    $bootType = "Remote"
                }

                $bootDevice = $device.Device
                $bootModel = $device.Model
                $bootVendor = $device.VEndor
                $bootSize = $device.Size
                $bootIsSAS = $TextInfo.ToTitleCase($device.IsSAS)
                $bootIsSSD = $TextInfo.ToTitleCase($device.IsSSD)
                $bootIsUSB = $TextInfo.ToTitleCase($device.IsUSB)
            }
        }

        # Pure Stateless (e.g. No USB or Disk for boot)
        if ($networkBoot -and $foundBootDevice -eq $false) {
            $bootModel = "N/A"
            $bootVendor = "N/A"
            $bootSize = "N/A"
            $bootIsSAS = "N/A"
            $bootIsSSD = "N/A"
            $bootIsUSB = "N/A"
        }

        $tmp = [PSCustomObject]@{
            Host = $($VMHost.ExtensionData.Name);
            Device = $bootDevice;
            BootType = $bootType;
            Vendor = $bootVendor;
            Model = $bootModel;
            SizeMB = $bootSize;
            IsSAS = $bootIsSAS;
            IsSSD = $bootIsSSD;
            IsUSB = $bootIsUSB;
        }
        $results += $tmp
        $results
    }

    function Get-ScsiDeviceDetail {
        <#
        .SYNOPSIS
        Helper function to return Scsi device information for a specific host and a specific datastore.
        .PARAMETER VMHosts
        This parameter accepts a list of host objects returned from the Get-VMHost cmdlet
        .PARAMETER VMHostMoRef
        This parameter specifies, by MoRef Id, the specific host of interest from with the $VMHosts array.
        .PARAMETER DatastoreDiskName
        This parameter specifies, by disk name, the specific datastore of interest.
        .EXAMPLE
        $VMHosts = Get-VMHost
        Get-ScsiDeviceDetail -AllVMHosts $VMHosts -VMHostMoRef 'HostSystem-host-131' -DatastoreDiskName 'naa.6005076801810082480000000001d9fe'
        DisplayName      : IBM Fibre Channel Disk (naa.6005076801810082480000000001d9fe)
        Ssd              : False
        LocalDisk        : False
        CanonicalName    : naa.6005076801810082480000000001d9fe
        Vendor           : IBM
        Model            : 2145
        Multipath Policy : Round Robin
        CapacityGB       : 512
        .NOTES
        Author: Ryan Kowalewski
    #>

        [CmdLetBinding()]
        param (
            [Parameter(Mandatory = $true)]
            $VMHosts,
            [Parameter(Mandatory = $true)]
            $VMHostMoRef,
            [Parameter(Mandatory = $true)]
            $DatastoreDiskName
        )

        $VMHostObj = $VMHosts | Where-Object { $_.Id -eq $VMHostMoRef }
        $ScsiDisk = $VMHostObj.ExtensionData.Config.StorageDevice.ScsiLun | Where-Object {
            $_.CanonicalName -eq $DatastoreDiskName
        }
        $Multipath = $VMHostObj.ExtensionData.Config.StorageDevice.MultipathInfo.Lun | Where-Object {
            $_.Lun -eq $ScsiDisk.Key
        }
        $CapacityGB = [math]::Round((($ScsiDisk.Capacity.BlockSize * $ScsiDisk.Capacity.Block) / 1024 / 1024 / 1024), 2)

        [PSCustomObject]@{
            'DisplayName' = $ScsiDisk.DisplayName
            'Ssd' = $ScsiDisk.Ssd
            'LocalDisk' = $ScsiDisk.LocalDisk
            'CanonicalName' = $ScsiDisk.CanonicalName
            'Vendor' = $ScsiDisk.Vendor
            'Model' = $ScsiDisk.Model
            'MultipathPolicy' = Switch ($Multipath.Policy.Policy) {
                'VMW_PSP_RR' { 'Round Robin' }
                'VMW_PSP_FIXED' { 'Fixed' }
                'VMW_PSP_MRU' { 'Most Recently Used' }
                default { $Multipath.Policy.Policy }
            }
            'Paths' = ($Multipath.Path).Count
            'CapacityGB' = $CapacityGB
        }
    }

    Function Get-PciDeviceDetail {
        <#
    .SYNOPSIS
    Helper function to return PCI Devices Drivers & Firmware information for a specific host.
    .PARAMETER Server
    vCenter VISession object.
    .PARAMETER esxcli
    Esxcli session object associated to the host.
    .EXAMPLE
    $Credentials = Get-Credential
    $Server = Connect-VIServer -Server vcenter01.example.com -Credentials $Credentials
    $VMHost = Get-VMHost -Server $Server -Name esx01.example.com
    $esxcli = Get-EsxCli -Server $Server -VMHost $VMHost -V2
    Get-PciDeviceDetail -Server $vCenter -esxcli $esxcli
    VMkernel Name    : vmhba0
    Device Name      : Sunrise Point-LP AHCI Controller
    Driver           : vmw_ahci
    Driver Version   : 1.0.0-34vmw.650.0.14.5146846
    Firmware Version : NA
    VIB Name         : vmw-ahci
    VIB Version      : 1.0.0-34vmw.650.0.14.5146846
    .NOTES
    Author: Erwan Quelin heavily based on the work of the vDocumentation team - https://github.com/arielsanchezmora/vDocumentation/blob/master/powershell/vDocumentation/Public/Get-ESXIODevice.ps1
    #>
        [CmdletBinding()]
        Param (
            [Parameter(Mandatory = $true)]
            $Server,
            [Parameter(Mandatory = $true)]
            $esxcli
        )
        Begin { }
    
        Process {
            # Set default results
            $firmwareVersion = "N/A"
            $vibName = "N/A"
            $driverVib = @{
                Name = "N/A"
                Version = "N/A"
            }
            $pciDevices = $esxcli.hardware.pci.list.Invoke() | Where-Object { $_.VMkernelName -like "vmhba*" -or $_.VMkernelName -like "vmnic*" -or $_.VMkernelName -like "vmgfx*" } | Sort-Object -Property VMkernelName 
            $nicList = $esxcli.network.nic.list.Invoke() | Sort-Object Name
            #$fcoeAdapterList = $esxcli.fcoe.adapter.list.Invoke().PhysicalNIC # Get list of vmnics used for FCoE, because we don't want those vmnics here.
            foreach ($pciDevice in $pciDevices) {
                $driverVersion = $esxcli.system.module.get.Invoke(@{module = $pciDevice.ModuleName }) | Select-Object -ExpandProperty Version
                # Get NIC Firmware version
                #if (($pciDevice.VMkernelName -like 'vmnic*') -and ($fcoeAdapterList -notcontains $pciDevice.VMkernelName) -and ($nicList.Name -contains $pciDevice.VMkernelName) ) {
                if (($pciDevice.VMkernelName -like 'vmnic*') -and ($nicList.Name -contains $pciDevice.VMkernelName) ) {   
                    $vmnicDetail = $esxcli.network.nic.get.Invoke(@{nicname = $pciDevice.VMkernelName })
                    $firmwareVersion = $vmnicDetail.DriverInfo.FirmwareVersion
                    # Get NIC driver VIB package version
                    $driverVib = $esxcli.software.vib.list.Invoke() | Select-Object -Property Name, Version | Where-Object { $_.Name -eq $vmnicDetail.DriverInfo.Driver -or $_.Name -eq "net-" + $vmnicDetail.DriverInfo.Driver -or $_.Name -eq "net55-" + $vmnicDetail.DriverInfo.Driver }
                    <#
                    If HP Smart Array vmhba* (scsi-hpsa driver) then get Firmware version
                    else skip if VMkernnel is vmhba*. Can't get HBA Firmware from 
                    Powercli at the moment only through SSH or using Putty Plink+PowerCli.
                    #>
                } elseif ($pciDevice.VMkernelName -like 'vmhba*') {
                    if ($pciDevice.DeviceName -match "smart array") {
                        $hpsa = $vmhost.ExtensionData.Runtime.HealthSystemRuntime.SystemHealthInfo.NumericSensorInfo | Where-Object { $_.Name -match "HP Smart Array" }
                        if ($hpsa) {
                            $firmwareVersion = (($hpsa.Name -split "firmware")[1]).Trim()
                        }
                    }
                    # Get HBA driver VIB package version
                    $vibName = $pciDevice.ModuleName -replace "_", "-"
                    $driverVib = $esxcli.software.vib.list.Invoke() | Select-Object -Property Name, Version | Where-Object { $_.Name -eq "scsi-" + $VibName -or $_.Name -eq "sata-" + $VibName -or $_.Name -eq $VibName }
                }
                # Output collected data
                [PSCustomObject]@{
                    'VMkernel Name' = $pciDevice.VMkernelName
                    'Device Name' = $pciDevice.DeviceName
                    'Driver' = $pciDevice.ModuleName
                    'Driver Version' = $driverVersion
                    'Firmware Version' = $firmwareVersion
                    'VIB Name' = $driverVib.Name
                    'VIB Version' = $driverVib.Version
                }
            } 
        }
        End { }
    }
    #endregion Script Functions

    #region Script Body
    #---------------------------------------------------------------------------------------------#
    #                                         SCRIPT BODY                                         #
    #---------------------------------------------------------------------------------------------#
    # Connect to ESXi Server using supplied credentials
    foreach ($VIServer in $Target) { 
        try {
            $ESXi = Connect-VIServer $VIServer -Credential $Credential -ErrorAction Stop
        } catch {
            Write-Error $_
        }
        #region Generate ESXi report
        if ($ESXi) {
            # Create a lookup hashtable to quickly link VM MoRefs to Names
            # Exclude VMware Site Recovery Manager placeholder VMs
            $VMs = Get-VM -Server $ESXi | Where-Object {
                $_.ExtensionData.Config.ManagedBy.ExtensionKey -notlike 'com.vmware.vcDr*'
            } | Sort-Object Name
            $VMLookup = @{ }
            foreach ($VM in $VMs) {
                $VMLookup.($VM.Id) = $VM.Name
            }

            # Create a lookup hashtable to link Host MoRefs to Names
            # Exclude VMware HCX hosts and ESX/ESXi versions prior to vSphere 5.0 from VMHost lookup
            $VMHosts = Get-VMHost -Server $ESXi | Where-Object { $_.Model -notlike "*VMware Mobility Platform" -and $_.Version -gt 5 } | Sort-Object Name
            $VMHostLookup = @{ }
            foreach ($VMHost in $VMHosts) {
                $VMHostLookup.($VMHost.Id) = $VMHost.ExtensionData.Name
            }

            # Create a lookup hashtable to link Datastore MoRefs to Names
            $Datastores = Get-Datastore -Server $ESXi | Where-Object { ($_.State -eq 'Available') -and ($_.CapacityGB -gt 0) } | Sort-Object Name
            $DatastoreLookup = @{ }
            foreach ($Datastore in $Datastores) {
                $DatastoreLookup.($Datastore.Id) = $Datastore.Name
            }

            # Create a lookup hashtable to link VDS Portgroups MoRefs to Names
            $VDPortGroups = Get-VDPortgroup -Server $ESXi | Sort-Object Name
            $VDPortGroupLookup = @{ }
            foreach ($VDPortGroup in $VDPortGroups) {
                $VDPortGroupLookup.($VDPortGroup.Key) = $VDPortGroup.Name
            }

            if ($InfoLevel.VMHost -ge 1) {
                if ($VMHosts) {
                    #region Hosts Section
                    Section -Style Heading1 $($VMHost.ExtensionData.Name) {
                        Paragraph "The following sections detail the configuration of VMware ESXi host $($VMHost.ExtensionData.Name)."
                        #region ESXi Host Detailed Information
                        foreach ($VMHost in ($VMHosts | Where-Object { $_.ConnectionState -eq 'Connected' -or $_.ConnectionState -eq 'Maintenance' })) {        
                            ### TODO: Host Certificate, Swap File Location
                            #region ESXi Host Hardware Section
                            Section -Style Heading2 'Hardware' {
                                Paragraph "The following section details the host hardware configuration for $($VMHost.ExtensionData.Name)."
                                BlankLine

                                #region ESXi Host Specifications
                                $VMHostUptime = Get-Uptime -VMHost $VMHost
                                $esxcli = Get-EsxCli -VMHost $VMHost -V2
                                $VMHostHardware = Get-VMHostHardware -VMHost $VMHost
                                $VMHostLicense = Get-VMHostLicense -VMHost $VMHost
                                $ScratchLocation = Get-AdvancedSetting -Entity $VMHost | Where-Object { $_.Name -eq 'ScratchConfig.CurrentScratchLocation' }
                                $VMHostDetail = [PSCustomObject]@{
                                    'Host' = $VMHost.ExtensionData.Name
                                    'Connection State' = Switch ($VMHost.ConnectionState) {
                                        'NotResponding' { 'Not Responding' }
                                        default { $VMHost.ConnectionState }
                                    }
                                    'ID' = $VMHost.Id
                                    'Manufacturer' = $VMHost.Manufacturer
                                    'Model' = $VMHost.Model
                                    'Serial Number' = $VMHostHardware.SerialNumber 
                                    'Asset Tag' = Switch ($VMHostHardware.AssetTag) {
                                        '' { 'Unknown' }
                                        default { $VMHostHardware.AssetTag }
                                    }
                                    'Processor Type' = $VMHost.Processortype
                                    'HyperThreading' = Switch ($VMHost.HyperthreadingActive) {
                                        $true { 'Enabled' }
                                        $false { 'Disabled' }
                                    }
                                    'Number of CPU Sockets' = $VMHost.ExtensionData.Hardware.CpuInfo.NumCpuPackages 
                                    'Number of CPU Cores' = $VMHost.ExtensionData.Hardware.CpuInfo.NumCpuCores 
                                    'Number of CPU Threads' = $VMHost.ExtensionData.Hardware.CpuInfo.NumCpuThreads
                                    'CPU Total / Used' = "$([math]::Round(($VMHost.CpuTotalMhz) / 1000, 2)) GHz / $([math]::Round(($VMHost.CpuUsageMhz) / 1000, 2)) GHz"
                                    'Memory Total / Used' = "$([math]::Round($VMHost.MemoryTotalGB, 2)) GB / $([math]::Round($VMHost.MemoryUsageGB, 2)) GB"
                                    'NUMA Nodes' = $VMHost.ExtensionData.Hardware.NumaInfo.NumNodes 
                                    'Number of NICs' = $VMHostHardware.NicCount 
                                    'Number of Datastores' = $VMHost.ExtensionData.Datastore.Count 
                                    'Number of VMs' = $VMHost.ExtensionData.VM.Count 
                                    'Power Management Policy' = $VMHost.ExtensionData.Hardware.CpuPowerManagementInfo.CurrentPolicy 
                                    'Scratch Location' = $ScratchLocation.Value 
                                    'Bios Version' = $VMHost.ExtensionData.Hardware.BiosInfo.BiosVersion 
                                    'Bios Release Date' = $VMHost.ExtensionData.Hardware.BiosInfo.ReleaseDate 
                                    'ESXi Version' = $VMHost.Version 
                                    'ESXi Build' = $VMHost.build 
                                    'Product' = $VMHostLicense.Product -join ', '
                                    'License Key' = $VMHostLicense.LicenseKey
                                    'License Expiration' = $VMHostLicense.Expiration 
                                    'Boot Time' = ($VMHost.ExtensionData.Runtime.Boottime).ToLocalTime()
                                    'Uptime Days' = $VMHostUptime.UptimeDays
                                }
                                if ($Healthcheck.VMHost.ConnectionState) {
                                    $VMHostDetail | Where-Object { $_.'Connection State' -eq 'Maintenance' } | Set-Style -Style Warning -Property 'Connection State'
                                }
                                if ($Healthcheck.VMHost.HyperThreading) {
                                    $VMHostDetail | Where-Object { $_.'HyperThreading' -eq 'Disabled' } | Set-Style -Style Warning -Property 'Disabled'
                                }
                                if ($Healthcheck.VMHost.Licensing) {
                                    $VMHostDetail | Where-Object { $_.'Product' -like '*Evaluation*' } | Set-Style -Style Warning -Property 'Product'
                                    $VMHostDetail | Where-Object { $_.'License Key' -like '*-00000-00000' } | Set-Style -Style Warning -Property 'License Key'
                                    $VMHostDetail | Where-Object { $_.'License Expiration' -eq 'Expired' } | Set-Style -Style Critical -Property 'License Expiration'
                                }
                                if ($Healthcheck.VMHost.ScratchLocation) {
                                    $VMHostDetail | Where-Object { $_.'Scratch Location' -eq '/tmp/scratch' } | Set-Style -Style Warning -Property 'Scratch Location'
                                }
                                if ($Healthcheck.VMHost.UpTimeDays) {
                                    $VMHostDetail | Where-Object { $_.'Uptime Days' -ge 275 -and $_.'Uptime Days' -lt 365 } | Set-Style -Style Warning -Property 'Uptime Days'
                                    $VMHostDetail | Where-Object { $_.'Uptime Days' -ge 365 } | Set-Style -Style Critical -Property 'Uptime Days'
                                }
                                $VMHostDetail | Table -Name "$($VMHost.ExtensionData.Name) ESXi Host Detailed Information" -List -ColumnWidths 50, 50 
                                #endregion ESXi Host Specifications

                                #region ESXi Host Boot Device
                                Section -Style Heading3 'Boot Device' {
                                    $ESXiBootDevice = Get-ESXiBootDevice -VMHost $VMHost
                                    $VMHostBootDevice = [PSCustomObject]@{
                                        'Host' = $ESXiBootDevice.Host
                                        'Device' = $ESXiBootDevice.Device
                                        'Boot Type' = $ESXiBootDevice.BootType
                                        'Vendor' = $ESXiBootDevice.Vendor
                                        'Model' = $ESXiBootDevice.Model
                                        'Size' = "$([math]::Round($ESXiBootDevice.SizeMB / 1024, 2)) GB"
                                        'Is SAS' = $ESXiBootDevice.IsSAS
                                        'Is SSD' = $ESXiBootDevice.IsSSD
                                        'Is USB' = $ESXiBootDevice.IsUSB
                                    }
                                    $VMHostBootDevice | Table -Name "$($VMHost.ExtensionData.Name) Boot Device" -List -ColumnWidths 50, 50 
                                }
                                #endregion ESXi Host Boot Devices

                                #region ESXi Host PCI Devices
                                Section -Style Heading3 'PCI Devices' {
                                    $PciHardwareDevices = $esxcli.hardware.pci.list.Invoke() | Where-Object { $_.VMkernelName -like "vmhba*" -OR $_.VMkernelName -like "vmnic*" -OR $_.VMkernelName -like "vmgfx*" } 
                                    $VMHostPciDevices = foreach ($PciHardwareDevice in $PciHardwareDevices) {
                                        [PSCustomObject]@{
                                            'VMkernel Name' = $PciHardwareDevice.VMkernelName 
                                            'PCI Address' = $PciHardwareDevice.Address 
                                            'Device Class' = $PciHardwareDevice.DeviceClassName 
                                            'Device Name' = $PciHardwareDevice.DeviceName 
                                            'Vendor Name' = $PciHardwareDevice.VendorName 
                                            'Slot Description' = $PciHardwareDevice.SlotDescription
                                        }
                                    }
                                    $VMHostPciDevices | Sort-Object 'VMkernel Name' | Table -Name "$($VMHost.ExtensionData.Name) PCI Devices" 
                                }
                                #endregion ESXi Host PCI Devices
                    
                                #region ESXi Host PCI Devices Drivers & Firmware
                                Section -Style Heading3 'PCI Devices Drivers & Firmware' {
                                    $VMHostPciDevicesDetails = Get-PciDeviceDetail -Server $VMHost -esxcli $esxcli 
                                    $VMHostPciDevicesDetails | Sort-Object 'VMkernel Name' | Table -Name "$($VMHost.ExtensionData.Name) PCI Devices Drivers & Firmware" 
                                }
                                #endregion ESXi Host PCI Devices Drivers & Firmware
                                #>
                            }
                            #endregion ESXi Host Hardware Section

                            #region ESXi Host System Section
                            Section -Style Heading2 'System' {
                                Paragraph "The following section details the host system configuration for $($VMHost.ExtensionData.Name)."
                                #region ESXi Host Image Profile Information
                                Section -Style Heading3 'Image Profile' {
                                    $installdate = Get-InstallDate
                                    $esxcli = Get-EsxCli -VMHost $VMHost -V2
                                    $ImageProfile = $esxcli.software.profile.get.Invoke()
                                    $SecurityProfile = [PSCustomObject]@{
                                        'Image Profile' = $ImageProfile.Name
                                        'Vendor' = $ImageProfile.Vendor
                                        'Installation Date' = $InstallDate.InstallDate
                                    }
                                    $SecurityProfile | Table -Name "$($VMHost.ExtensionData.Name) Image Profile" -ColumnWidths 50, 25, 25 
                                }
                                #endregion ESXi Host Image Profile Information

                                #region ESXi Host Time Configuration
                                Section -Style Heading3 'Time Configuration' {
                                    $VMHostTimeSettings = [PSCustomObject]@{
                                        'Time Zone' = $VMHost.timezone
                                        'NTP Service' = Switch ((Get-VMHostService -VMHost $VMHost | Where-Object { $_.key -eq 'ntpd' }).Running) {
                                            $true { 'Running' }
                                            $false { 'Stopped' }
                                        }
                                        'NTP Server(s)' = (Get-VMHostNtpServer -VMHost $VMHost | Sort-Object) -join ', '
                                    }
                                    if ($Healthcheck.VMHost.NTP) {
                                        $VMHostTimeSettings | Where-Object { $_.'NTP Service' -eq 'Stopped' } | Set-Style -Style Critical -Property 'NTP Service'
                                    }
                                    $VMHostTimeSettings | Table -Name "$($VMHost.ExtensionData.Name) Time Configuration" -ColumnWidths 30, 30, 40
                                }
                                #endregion ESXi Host Time Configuration

                                #region ESXi Host Syslog Configuration
                                $SyslogConfig = $VMHost | Get-VMHostSysLogServer
                                if ($SyslogConfig) {
                                    Section -Style Heading3 'Syslog Configuration' {
                                        ### TODO: Syslog Rotate & Size, Log Directory (Adv Settings)
                                        $SyslogConfig = $SyslogConfig | Select-Object @{L = 'SysLog Server'; E = { $_.Host } }, Port
                                        $SyslogConfig | Table -Name "$($VMHost.ExtensionData.Name) Syslog Configuration" -ColumnWidths 50, 50 
                                    }
                                }
                                #endregion ESXi Host Syslog Configuration

                                #region ESXi Host Comprehensive Information Section
                                if ($InfoLevel.VMHost -ge 5) {
                                    #region ESXi Host Advanced System Settings
                                    Section -Style Heading3 'Advanced System Settings' {
                                        $AdvSettings = $VMHost | Get-AdvancedSetting | Select-Object Name, Value
                                        $AdvSettings | Sort-Object Name | Table -Name "$($VMHost.ExtensionData.Name) Advanced System Settings" -ColumnWidths 50, 50 
                                    }
                                    #endregion ESXi Host Advanced System Settings

                                    #region ESXi Host Software VIBs
                                    Section -Style Heading3 'Software VIBs' {
                                        $esxcli = Get-EsxCli -VMHost $VMHost -V2
                                        $VMHostVibs = $esxcli.software.vib.list.Invoke()
                                        $VMHostVibs = foreach ($VMHostVib in $VMHostVibs) {
                                            [PSCustomObject]@{
                                                'VIB' = $VMHostVib.Name
                                                'ID' = $VMHostVib.Id
                                                'Version' = $VMHostVib.Version
                                                'Acceptance Level' = $VMHostVib.AcceptanceLevel
                                                'Creation Date' = $VMHostVib.CreationDate
                                                'Install Date' = $VMHostVib.InstallDate
                                            }
                                        } 
                                        $VMHostVibs | Sort-Object 'Install Date' -Descending | Table -Name "$($VMHost.ExtensionData.Name) Software VIBs" -ColumnWidths 15, 25, 15, 15, 15, 15
                                    }
                                    #endregion ESXi Host Software VIBs
                                }
                                #endregion ESXi Host Comprehensive Information Section
                            }
                            #endregion ESXi Host System Section

                            #region ESXi Host Storage Section
                            Section -Style Heading2 'Storage' {
                                Paragraph "The following section details the host storage configuration for $($VMHost.ExtensionData.Name)."
                                
                                #region Datastore Section
                                # Currently there is no Datastore InfoLevel 1
                                if ($InfoLevel.Datastore -ge 2) {
                                    if ($Datastores) {
                                        Section -Style Heading3 'Datastores' {
                                            #region Datastore Infomative Information
                                            if ($InfoLevel.Datastore -eq 2) {
                                                $DatastoreInfo = foreach ($Datastore in $Datastores) {
                                                    [PSCustomObject]@{
                                                        'Datastore' = $Datastore.Name
                                                        'Type' = $Datastore.Type
                                                        'Version' = Switch ($Datastore.FileSystemVersion) {
                                                            $null { '--' }
                                                            default { $Datastore.FileSystemVersion }
                                                        }
                                                        '# of VMs' = $Datastore.ExtensionData.VM.Count
                                                        'Total Capacity GB' = [math]::Round($Datastore.CapacityGB, 2)
                                                        'Used Capacity GB' = [math]::Round((($Datastore.CapacityGB) - ($Datastore.FreeSpaceGB)), 2)
                                                        'Free Space GB' = [math]::Round($Datastore.FreeSpaceGB, 2)
                                                        '% Used' = [math]::Round((100 - (($Datastore.FreeSpaceGB) / ($Datastore.CapacityGB) * 100)), 2)
                                                    }
                                                }
                                                if ($Healthcheck.Datastore.CapacityUtilization) {
                                                    $DatastoreInfo | Where-Object { $_.'% Used' -ge 90 } | Set-Style -Style Critical -Property '% Used'
                                                    $DatastoreInfo | Where-Object { $_.'% Used' -ge 75 -and $_.'% Used' -lt 90 } | Set-Style -Style Warning -Property '% Used'
                                                }
                                                $DatastoreInfo | Sort-Object Datastore | Table -Name 'Datastore Information'
                                            }
                                            #endregion Datastore Informative Information

                                            #region Datastore Detailed Information
                                            if ($InfoLevel.Datastore -ge 3) {
                                                foreach ($Datastore in $Datastores) {
                                                    #region Datastore Section
                                                    Section -Style Heading4 $Datastore.Name {                                
                                                        $DatastoreDetail = [PSCustomObject]@{
                                                            'Datastore' = $Datastore.Name
                                                            'ID' = $Datastore.Id
                                                            'Type' = $Datastore.Type
                                                            'Version' = Switch ($Datastore.FileSystemVersion) {
                                                                $null { '--' }
                                                                default { $Datastore.FileSystemVersion }
                                                            }
                                                            'State' = $Datastore.State
                                                            'Number of VMs' = $Datastore.ExtensionData.VM.Count
                                                            'Storage I/O Control' = Switch ($Datastore.StorageIOControlEnabled) {
                                                                $true { 'Enabled' }
                                                                $false { 'Disabled' }
                                                            }
                                                            'Congestion Threshold' = Switch ($Datastore.CongestionThresholdMillisecond) {
                                                                $null { '--' }
                                                                default { "$($Datastore.CongestionThresholdMillisecond) ms" }
                                                            }
                                                            'Total Capacity' = "$([math]::Round($Datastore.CapacityGB, 2)) GB"
                                                            'Used Capacity' = "$([math]::Round((($Datastore.CapacityGB) - ($Datastore.FreeSpaceGB)), 2)) GB"
                                                            'Free Space' = "$([math]::Round($Datastore.FreeSpaceGB, 2)) GB"
                                                            '% Used' = [math]::Round((100 - (($Datastore.FreeSpaceGB) / ($Datastore.CapacityGB) * 100)), 2)
                                                        }
                                                        if ($Healthcheck.Datastore.CapacityUtilization) {
                                                            $DatastoreDetail | Where-Object { $_.'% Used' -ge 90 } | Set-Style -Style Critical -Property '% Used'
                                                            $DatastoreDetail | Where-Object { $_.'% Used' -ge 75 -and 
                                                                $_.'% Used' -lt 90 } | Set-Style -Style Warning -Property '% Used'
                                                        }
                        
                                                        #region Datastore Advanced Detailed Information
                                                        if ($InfoLevel.Datastore -ge 4) {
                                                            $MemberProps = @{
                                                                'InputObject' = $DatastoreDetail
                                                                'MemberType' = 'NoteProperty'
                                                            }
                                                            $DatastoreVMs = foreach ($DatastoreVM in $Datastore.ExtensionData.VM) {
                                                                $VMLookup."$($DatastoreVM.Type)-$($DatastoreVM.Value)"
                                                            }
                                                            Add-Member @MemberProps -Name 'Virtual Machines' -Value (($DatastoreVMs | Sort-Object) -join ', ')
                                                        }
                                                        #endregion Datastore Advanced Detailed Information

                                                        $DatastoreDetail | Sort-Object Datacenter, Datastore | Table -List -Name 'Datastore Specifications' -ColumnWidths 50, 50

                                                        # Get VMFS volumes. Ignore local SCSILuns.
                                                        if (($Datastore.Type -eq 'VMFS') -and ($Datastore.ExtensionData.Info.Vmfs.Local -eq $false)) {
                                                            #region SCSI LUN Information Section
                                                            Section -Style Heading4 'SCSI LUN Information' {
                                                                $ScsiLuns = foreach ($DatastoreHost in $Datastore.ExtensionData.Host.Key) {
                                                                    $DiskName = $Datastore.ExtensionData.Info.Vmfs.Extent.DiskName
                                                                    $ScsiDeviceDetailProps = @{
                                                                        'VMHosts' = $VMHosts
                                                                        'VMHostMoRef' = "$($DatastoreHost.Type)-$($DatastoreHost.Value)"
                                                                        'DatastoreDiskName' = $DiskName
                                                                    }
                                                                    $ScsiDeviceDetail = Get-ScsiDeviceDetail @ScsiDeviceDetailProps

                                                                    [PSCustomObject]@{
                                                                        'Host' = $VMHostLookup."$($DatastoreHost.Type)-$($DatastoreHost.Value)"
                                                                        'Canonical Name' = $DiskName
                                                                        'Capacity GB' = $ScsiDeviceDetail.CapacityGB
                                                                        'Vendor' = $ScsiDeviceDetail.Vendor
                                                                        'Model' = $ScsiDeviceDetail.Model
                                                                        'Is SSD' = $ScsiDeviceDetail.Ssd
                                                                        'Multipath Policy' = $ScsiDeviceDetail.MultipathPolicy
                                                                        'Paths' = $ScsiDeviceDetail.Paths
                                                                    }
                                                                }
                                                                $ScsiLuns | Sort-Object Host | Table -Name 'SCSI LUN Information'
                                                            }
                                                            #endregion SCSI LUN Information Section
                                                        }
                                                    }
                                                    #endregion Datastore Section
                                                }
                                            }
                                            #endregion Datastore Detailed Information
                                        }
                                    }
                                }
                                #endregion Datastore Section

                                #region ESXi Host Storage Adapter Information
                                $VMHostHbas = $VMHost | Get-VMHostHba | Sort-Object Device
                                if ($VMHostHbas) {
                                    #region ESXi Host Storage Adapters Section
                                    Section -Style Heading3 'Storage Adapters' {
                                        foreach ($VMHostHba in $VMHostHbas) {
                                            $Target = ((Get-View $VMHostHba.VMhost).Config.StorageDevice.ScsiTopology.Adapter | Where-Object { $_.Adapter -eq $VMHostHba.Key }).Target
                                            $LUNs = Get-ScsiLun -Hba $VMHostHba -LunType "disk" -ErrorAction SilentlyContinue
                                            $Paths = ($Target | foreach { $_.Lun.Count } | Measure-Object -Sum)
                                            Section -Style Heading4 "$($VMHostHba.Device)" {
                                                $VMHostStorageAdapter = [PSCustomObject]@{
                                                    'Adapter' = $VMHostHba.Device
                                                    'Type' = Switch ($VMHostHba.Type) {
                                                        'FibreChannel' { 'Fibre Channel' }
                                                        'IScsi' { 'iSCSI' }
                                                        'ParallelScsi' { 'Parallel SCSI' }
                                                        default { $TextInfo.ToTitleCase($VMHostHba.Type) }
                                                    }
                                                    'Model' = $VMHostHba.Model
                                                    'Status' = $TextInfo.ToTitleCase($VMHostHba.Status)
                                                    'Targets' = $Target.Count
                                                    'Devices' = $LUNs.Count
                                                    'Paths' = $Paths.Sum
                                                }
                                                $MemberProps = @{
                                                    'InputObject' = $VMHostStorageAdapter
                                                    'MemberType' = 'NoteProperty'
                                                }
                                                if ($VMHostStorageAdapter.Type -eq 'iSCSI') {
                                                    $iScsiAuthenticationMethod = Switch ($VMHostHba.ExtensionData.AuthenticationProperties.ChapAuthenticationType) {
                                                        'chapProhibited' { 'None' }
                                                        'chapPreferred' { 'Use unidirectional CHAP unless prohibited by target' }
                                                        'chapDiscouraged' { 'Use unidirectional CHAP if required by target' }
                                                        'chapRequired' { 
                                                            Switch ($VMHostHba.ExtensionData.AuthenticationProperties.MutualChapAuthenticationType) {
                                                                'chapProhibited' { 'Use unidirectional CHAP' }
                                                                'chapRequired' { 'Use bidirectional CHAP' }
                                                            } 
                                                        }
                                                        default { $VMHostHba.ExtensionData.AuthenticationProperties.ChapAuthenticationType }
                                                    }
                                                    Add-Member @MemberProps -Name 'iSCSI Name' -Value $VMHostHba.IScsiName
                                                    if ($VMHostHba.IScsiAlias) {
                                                        Add-Member @MemberProps -Name 'iSCSI Alias' -Value $VMHostHba.IScsiAlias
                                                    } else {
                                                        Add-Member @MemberProps -Name 'iSCSI Alias' -Value '--'
                                                    }
                                                    if ($VMHostHba.CurrentSpeedMb) {
                                                        Add-Member @MemberProps -Name 'Speed' -Value "$($VMHostHba.CurrentSpeedMb) Mb"
                                                    } else {
                                                        Add-Member @MemberProps -Name 'Speed' -Value '--'
                                                    }
                                                    if ($VMHostHba.ExtensionData.ConfiguredSendTarget) {
                                                        Add-Member @MemberProps -Name 'Dynamic Discovery' -Value (($VMHostHba.ExtensionData.ConfiguredSendTarget | ForEach-Object { "$($_.Address)" + ":" + "$($_.Port)" }) -join [Environment]::NewLine)
                                                    } else {
                                                        Add-Member @MemberProps -Name 'Dynamic Discovery' -Value '--'
                                                    }
                                                    if ($VMHostHba.ExtensionData.ConfiguredStaticTarget) {
                                                        Add-Member @MemberProps -Name 'Static Discovery' -Value (($VMHostHba.ExtensionData.ConfiguredStaticTarget | ForEach-Object { "$($_.Address)" + ":" + "$($_.Port)" + "  " + "$($_.IScsiName)" }) -join [Environment]::NewLine)
                                                    } else {
                                                        Add-Member @MemberProps -Name 'Static Discovery' -Value '--'
                                                    }
                                                    if ($iScsiAuthenticationMethod -eq 'None') {
                                                        Add-Member @MemberProps -Name 'Authentication Method' -Value $iScsiAuthenticationMethod
                                                    } elseif ($iScsiAuthenticationMethod -eq 'Use bidirectional CHAP') {
                                                        Add-Member @MemberProps -Name 'Authentication Method' -Value $iScsiAuthenticationMethod
                                                        Add-Member @MemberProps -Name 'Outgoing CHAP Name' -Value $VMHostHba.ExtensionData.AuthenticationProperties.ChapName
                                                        Add-Member @MemberProps -Name 'Incoming CHAP Name' -Value $VMHostHba.ExtensionData.AuthenticationProperties.MutualChapName
                                                    } else {
                                                        Add-Member @MemberProps -Name 'Authentication Method' -Value $iScsiAuthenticationMethod
                                                        Add-Member @MemberProps -Name 'Outgoing CHAP Name' -Value $VMHostHba.ExtensionData.AuthenticationProperties.ChapName
                                                    }
                                                    if ($InfoLevel.VMHost -eq 4) {
                                                        Add-Member @MemberProps -Name 'Advanced Options' -Value (($VMHostHba.ExtensionData.AdvancedOptions | ForEach-Object { "$($_.Key) = $($_.Value)" }) -join [Environment]::NewLine)
                                                    }
                                                }
                                                if ($VMHostStorageAdapter.Type -eq 'Fibre Channel') {
                                                    Add-Member @MemberProps -Name 'Node WWN' -Value (([String]::Format("{0:X}", $VMHostHba.NodeWorldWideName) -split "(\w{2})" | Where-Object { $_ -ne "" }) -join ":")
                                                    Add-Member @MemberProps -Name 'Port WWN' -Value (([String]::Format("{0:X}", $VMHostHba.PortWorldWideName) -split "(\w{2})" | Where-Object { $_ -ne "" }) -join ":")
                                                    Add-Member @MemberProps -Name 'Speed' -Value $VMHostHba.Speed
                                                }
                                                if ($Healthcheck.VMHost.StorageAdapter) {
                                                    $VMHostStorageAdapter | Where-Object { $_.'Status' -ne 'Online' } | Set-Style -Style Warning -Property 'Status'
                                                    $VMHostStorageAdapter | Where-Object { $_.'Status' -eq 'Offline' } | Set-Style -Style Critical -Property 'Status'
                                                }
                                                $VMHostStorageAdapter | Table -List -Name "$($VMHost.ExtensionData.Name) storage adapter $($VMHostStorageAdapter.Adapter)" -ColumnWidths 25, 75
                                            }
                                        }
                                    }
                                    #endregion ESXi Host Storage Adapters Section
                                }
                                #endregion ESXi Host Storage Adapter Information
                            }
                            #endregion ESXi Host Storage Section

                            #region ESXi Host Network Section
                            Section -Style Heading2 'Network' {
                                Paragraph "The following section details the host network configuration for $($VMHost.ExtensionData.Name)."
                                BlankLine
                                #region ESXi Host Network Configuration
                                $VMHostNetwork = $VMHost.ExtensionData.Config.Network
                                $VMHostVirtualSwitch = @()
                                $VMHostVss = foreach ($vSwitch in $VMHost.ExtensionData.Config.Network.Vswitch) {
                                    $VMHostVirtualSwitch += $vSwitch.Name
                                }
                                $VMHostDvs = foreach ($dvSwitch in $VMHost.ExtensionData.Config.Network.ProxySwitch) {
                                    $VMHostVirtualSwitch += $dvSwitch.DvsName
                                }
                                $VMHostNetworkDetail = [PSCustomObject]@{
                                    'Host' = $($VMHost.ExtensionData.Name)
                                    'Virtual Switches' = ($VMHostVirtualSwitch | Sort-Object) -join ', '
                                    'VMkernel Adapters' = ($VMHostNetwork.Vnic.Device | Sort-Object) -join ', '
                                    'Physical Adapters' = ($VMHostNetwork.Pnic.Device | Sort-Object) -join ', '
                                    'VMkernel Gateway' = $VMHostNetwork.IpRouteConfig.DefaultGateway
                                    'IPv6' = Switch ($VMHostNetwork.IPv6Enabled) {
                                        $true { 'Enabled' }
                                        $false { 'Disabled' }
                                    }
                                    'VMkernel IPv6 Gateway' = Switch ($VMHostNetwork.IpRouteConfig.IpV6DefaultGateway) {
                                        $null { '--' }
                                        default { $VMHostNetwork.IpRouteConfig.IpV6DefaultGateway }
                                    }
                                    'DNS Servers' = ($VMHostNetwork.DnsConfig.Address | Sort-Object) -join ', ' 
                                    'Host Name' = $VMHostNetwork.DnsConfig.HostName
                                    'Domain Name' = $VMHostNetwork.DnsConfig.DomainName 
                                    'Search Domain' = ($VMHostNetwork.DnsConfig.SearchDomain | Sort-Object) -join ', '
                                }
                                if ($Healthcheck.VMHost.IPv6) {
                                    $VMHostNetworkDetail | Where-Object { $_.'IPv6' -eq $false } | Set-Style -Style Warning -Property 'IPv6'
                                }
                                $VMHostNetworkDetail | Table -Name "$($VMHost.ExtensionData.Name) Network Configuration" -List -ColumnWidths 50, 50
                                #endregion ESXi Host Network Configuration

                                #region ESXi Host Physical Adapters
                                Section -Style Heading3 'Physical Adapters' {
                                    $PhysicalNetAdapters = $VMHost.ExtensionData.Config.Network.Pnic | Sort-Object Device
                                    $VMHostPhysicalNetAdapters = foreach ($PhysicalNetAdapter in $PhysicalNetAdapters) {
                                        [PSCustomObject]@{
                                            'Adapter' = $PhysicalNetAdapter.Device
                                            'Status' = Switch ($PhysicalNetAdapter.Linkspeed) {
                                                $null { 'Disconnected' }
                                                default { 'Connected' }
                                            }
                                            'Virtual Switch' = $(
                                                if ($VMHost.ExtensionData.Config.Network.Vswitch.Pnic -contains $PhysicalNetAdapter.Key) {
                                                    ($VMHost.ExtensionData.Config.Network.Vswitch | Where-Object { $_.Pnic -eq $PhysicalNetAdapter.Key }).Name
                                                } elseif ($VMHost.ExtensionData.Config.Network.ProxySwitch.Pnic -contains $PhysicalNetAdapter.Key) {
                                                    ($VMHost.ExtensionData.Config.Network.ProxySwitch | Where-Object { $_.Pnic -eq $PhysicalNetAdapter.Key }).DvsName
                                                } else {
                                                    '--'
                                                }
                                            )
                                            'MAC Address' = $PhysicalNetAdapter.Mac
                                            'Actual Speed, Duplex' = Switch ($PhysicalNetAdapter.LinkSpeed.SpeedMb) {
                                                $null { 'Down' }
                                                default {
                                                    if ($PhysicalNetAdapter.LinkSpeed.Duplex) {
                                                        "$($PhysicalNetAdapter.LinkSpeed.SpeedMb) Mbps, Full Duplex"
                                                    } else {
                                                        'Auto negotiate'
                                                    }
                                                }
                                            }
                                            'Configured Speed, Duplex' = Switch ($PhysicalNetAdapter.Spec.LinkSpeed) {
                                                $null { 'Auto negotiate' }
                                                default {
                                                    if ($PhysicalNetAdapter.Spec.LinkSpeed.Duplex) {
                                                        "$($PhysicalNetAdapter.Spec.LinkSpeed.SpeedMb) Mbps, Full Duplex"
                                                    } else {
                                                        "$($PhysicalNetAdapter.Spec.LinkSpeed.SpeedMb) Mbps"
                                                    }
                                                }
                                            }
                                            'Wake on LAN' = Switch ($PhysicalNetAdapter.WakeOnLanSupported) {
                                                $true { 'Supported' }
                                                $false { 'Not Supported' }
                                            }
                                        }
                                    }
                                    if ($Healthcheck.VMHost.NetworkAdapter) {
                                        $VMHostPhysicalNetAdapters | Where-Object { $_.'Status' -ne 'Connected' } | Set-Style -Style Critical -Property 'Status'
                                        $VMHostPhysicalNetAdapters | Where-Object { $_.'Actual Speed, Duplex' -eq 'Down' } | Set-Style -Style Critical -Property 'Actual Speed, Duplex'
                                    }
                                    if ($InfoLevel.VMHost -ge 4) {
                                        foreach ($VMHostPhysicalNetAdapter in $VMHostPhysicalNetAdapters) {
                                            Section -Style Heading4 "$($VMHostPhysicalNetAdapter.Adapter)" {
                                                $VMHostPhysicalNetAdapter | Table -List -Name "$($VMHost.ExtensionData.Name) Physical Adapter $($VMHostPhysicalNetAdapter.Adapter)" -ColumnWidths 50, 50
                                            }
                                        }
                                    } else {
                                        $VMHostPhysicalNetAdapters | Table -Name "$($VMHost.ExtensionData.Name) Physical Adapters"
                                    }
                                }
                                #endregion ESXi Host Physical Adapters
                    
                                #region ESXi Host Cisco Discovery Protocol
                                $VMHostNetworkAdapterCDP = $VMHost | Get-VMHostNetworkAdapterCDP | Where-Object { $_.Status -eq 'Connected' } | Sort-Object Device
                                if ($VMHostNetworkAdapterCDP) {
                                    Section -Style Heading3 'Cisco Discovery Protocol' {
                                        if ($InfoLevel.VMHost -ge 4) {
                                            foreach ($VMHostNetworkAdapter in $VMHostNetworkAdapterCDP) {
                                                Section -Style Heading4 "$($VMHostNetworkAdapter.Device)" {
                                                    $VMHostCDP = [PSCustomObject]@{
                                                        'Status' = $VMHostNetworkAdapter.Status
                                                        'System Name' = $VMHostNetworkAdapter.SystemName
                                                        'Hardware Platform' = $VMHostNetworkAdapter.HardwarePlatform
                                                        'Switch ID' = $VMHostNetworkAdapter.SwitchId
                                                        'Software Version' = $VMHostNetworkAdapter.SoftwareVersion
                                                        'Management Address' = $VMHostNetworkAdapter.ManagementAddress
                                                        'Address' = $VMHostNetworkAdapter.Address
                                                        'Port ID' = $VMHostNetworkAdapter.PortId
                                                        'VLAN' = $VMHostNetworkAdapter.Vlan
                                                        'MTU' = $VMHostNetworkAdapter.Mtu
                                                    }
                                                    $VMHostCDP | Table -List -Name "$($VMHost.ExtensionData.Name) Network Adapter $($VMHostNetworkAdapter.Device) CDP Information" -ColumnWidths 50, 50
                                                }
                                            }
                                        } else {
                                            $VMHostCDP = foreach ($VMHostNetworkAdapter in $VMHostNetworkAdapterCDP) {
                                                [PSCustomObject]@{
                                                    'Adapter' = $VMHostNetworkAdapter.Device
                                                    'Status' = $VMHostNetworkAdapter.Status
                                                    'Hardware Platform' = $VMHostNetworkAdapter.HardwarePlatform
                                                    'Switch ID' = $VMHostNetworkAdapter.SwitchId
                                                    'Address' = $VMHostNetworkAdapter.Address
                                                    'Port ID' = $VMHostNetworkAdapter.PortId
                                                }
                                            }
                                            $VMHostCDP | Table -Name "$($VMHost.ExtensionData.Name) Network Adapter CDP Information"
                                        }
                                    }
                                }
                                #endregion ESXi Host Cisco Discovery Protocol

                                #region ESXi Host VMkernel Adapaters
                                Section -Style Heading3 'VMkernel Adapters' {
                                    $VMkernelAdapters = $VMHost | Get-View | ForEach-Object -Process {
                                        $esx = $_
                                        $netSys = Get-View -Id $_.ConfigManager.NetworkSystem
                                        $vnicMgr = Get-View -Id $_.ConfigManager.VirtualNicManager
                                        $netSys.NetworkInfo.Vnic |
                                        ForEach-Object -Process {
                                            $device = $_.Device
                                            [PSCustomObject]@{
                                                'Adapter' = $_.Device
                                                'Port Group' = & {
                                                    if ($_.Spec.Portgroup) {
                                                        $script:pg = $_.Spec.Portgroup
                                                    } else {
                                                        $script:pg = Get-View -ViewType DistributedVirtualPortgroup -Property Name, Key -Filter @{'Key' = "$($_.Spec.DistributedVirtualPort.PortgroupKey)" } |
                                                        Select-Object -ExpandProperty Name
                                                    }
                                                    $script:pg
                                                }
                                                'Virtual Switch' = & { 
                                                    if ($_.Spec.Portgroup) {
                                                        (Get-VirtualPortGroup -Standard -Name $script:pg -VMHost $VMHost).VirtualSwitchName
                                                    } else {
                                                        (Get-VDPortgroup -Name $script:pg).VDSwitch.Name
                                                    }
                                                }
                                                'TCP/IP Stack' = Switch ($_.Spec.NetstackInstanceKey) {
                                                    'defaultTcpipStack' { 'Default' }
                                                    'vSphereProvisioning' { 'Provisioning' }
                                                    'vmotion' { 'vMotion' }
                                                    $null { 'Not Applicable' }
                                                    default { $_.Spec.NetstackInstanceKey }
                                                }
                                                'MTU' = $_.Spec.Mtu
                                                'MAC Address' = $_.Spec.Mac
                                                'DHCP' = Switch ($_.Spec.Ip.Dhcp) {
                                                    $true { 'Enabled' }
                                                    $false { 'Disabled' }
                                                }
                                                'IP Address' = $_.Spec.IP.IPAddress
                                                'Subnet Mask' = $_.Spec.IP.SubnetMask
                                                'Default Gateway' = Switch ($_.Spec.IpRouteSpec.IpRouteConfig.DefaultGateway) {
                                                    $null { '--' }
                                                    default { $_.Spec.IpRouteSpec.IpRouteConfig.DefaultGateway }
                                                }
                                                'vMotion' = Switch ((($vnicMgr.Info.NetConfig | where { $_.NicType -eq 'vmotion' }).SelectedVnic | % { $_ -match $device } ) -contains $true) {
                                                    $true { 'Enabled' }
                                                    $false { 'Disabled' }
                                                }
                                                'Provisioning' = Switch ((($vnicMgr.Info.NetConfig | where { $_.NicType -eq 'vSphereProvisioning' }).SelectedVnic | % { $_ -match $device } ) -contains $true) {
                                                    $true { 'Enabled' }
                                                    $false { 'Disabled' }
                                                }
                                                'FT Logging' = Switch ((($vnicMgr.Info.NetConfig | where { $_.NicType -eq 'faultToleranceLogging' }).SelectedVnic | % { $_ -match $device } ) -contains $true) {
                                                    $true { 'Enabled' }
                                                    $false { 'Disabled' }
                                                }
                                                'Management' = Switch ((($vnicMgr.Info.NetConfig | where { $_.NicType -eq 'management' }).SelectedVnic | % { $_ -match $device } ) -contains $true) {
                                                    $true { 'Enabled' }
                                                    $false { 'Disabled' }
                                                }
                                                'vSphere Replication' = Switch ((($vnicMgr.Info.NetConfig | where { $_.NicType -eq 'vSphereReplication' }).SelectedVnic | % { $_ -match $device } ) -contains $true) {
                                                    $true { 'Enabled' }
                                                    $false { 'Disabled' }
                                                }
                                                'vSphere Replication NFC' = Switch ((($vnicMgr.Info.NetConfig | where { $_.NicType -eq 'vSphereReplicationNFC' }).SelectedVnic | % { $_ -match $device } ) -contains $true) {
                                                    $true { 'Enabled' }
                                                    $false { 'Disabled' }
                                                }
                                                'vSAN' = Switch ((($vnicMgr.Info.NetConfig | where { $_.NicType -eq 'vsan' }).SelectedVnic | % { $_ -match $device } ) -contains $true) {
                                                    $true { 'Enabled' }
                                                    $false { 'Disabled' }
                                                }
                                                'vSAN Witness' = Switch ((($vnicMgr.Info.NetConfig | where { $_.NicType -eq 'vsanWitness' }).SelectedVnic | % { $_ -match $device } ) -contains $true) {
                                                    $true { 'Enabled' }
                                                    $false { 'Disabled' }
                                                }
                                            }
                                        }
                                    }
                                    foreach ($VMkernelAdapter in ($VMkernelAdapters | Sort-Object 'Adapter')) {
                                        Section -Style Heading4 "$($VMkernelAdapter.Adapter)" {
                                            $VMkernelAdapter | Table -List -Name "$($VMHost.ExtensionData.Name) VMkernel Adapter $($VMkernelAdapter.Adapter)" -ColumnWidths 50, 50
                                        }
                                    }
                                }
                                #endregion ESXi Host VMkernel Adapaters

                                #region ESXi Host Standard Virtual Switches
                                $VSSwitches = $VMHost | Get-VirtualSwitch -Standard | Sort-Object Name
                                if ($VSSwitches) {
                                    #region Section Standard Virtual Switches
                                    Section -Style Heading5 'Standard Virtual Switches' {
                                        Paragraph "The following section details the standard virtual switch configuration for $VMHost."
                                        BlankLine
                                        $VSSwitchNicTeaming = $VSSwitches | Get-NicTeamingPolicy
                                        #region ESXi Host Standard Virtual Switch Properties
                                        $VSSProperties = foreach ($VSSwitchNicTeam in $VSSwitchNicTeaming) {
                                            [PSCustomObject]@{
                                                'Virtual Switch' = $VSSwitchNicTeam.VirtualSwitch 
                                                'MTU' = $VSSwitchNicTeam.VirtualSwitch.Mtu 
                                                'Number of Ports' = $VSSwitchNicTeam.VirtualSwitch.NumPorts
                                                'Number of Ports Available' = $VSSwitchNicTeam.VirtualSwitch.NumPortsAvailable
                                            }
                                        }
                                        $VSSProperties | Table -Name "$VMHost Standard Virtual Switches"
                                        #endregion ESXi Host Standard Virtual Switch Properties
                                
                                        #region ESXi Host Virtual Switch Security Policy
                                        $VssSecurity = $VSSwitches | Get-SecurityPolicy
                                        if ($VssSecurity) {
                                            #region Virtual Switch Security Policy
                                            Section -Style Heading5 'Virtual Switch Security' {
                                                $VssSecurity = foreach ($VssSec in $VssSecurity) {
                                                    [PSCustomObject]@{
                                                        'Virtual Switch' = $VssSec.VirtualSwitch 
                                                        'Promiscuous Mode' = Switch ($VssSec.AllowPromiscuous) {
                                                            $true { 'Accept' }
                                                            $false { 'Reject' }
                                                        }
                                                        'MAC Address Changes' = Switch ($VssSec.MacChanges) {
                                                            $true { 'Accept' }
                                                            $false { 'Reject' }
                                                        } 
                                                        'Forged Transmits' = Switch ($VssSec.ForgedTransmits) {
                                                            $true { 'Accept' }
                                                            $false { 'Reject' }
                                                        } 
                                                    }
                                                }
                                                $VssSecurity | Sort-Object 'Virtual Switch' | Table -Name "$VMHost Virtual Switch Security Policy"
                                            }
                                            #endregion Virtual Switch Security Policy
                                        }
                                        #endregion ESXi Host Virtual Switch Security Policy 

                                        #region ESXi Host Virtual Switch Traffic Shaping Policy
                                        Section -Style Heading5 'Virtual Switch Traffic Shaping' {
                                            $VssTrafficShapingPolicy = foreach ($VSSwitch in $VSSwitches) {
                                                [PSCustomObject]@{
                                                    'Virtual Switch' = $VSSwitch.Name
                                                    'Status' = Switch ($VSSwitch.ExtensionData.Spec.Policy.ShapingPolicy.Enabled) {
                                                        $True { 'Enabled' }
                                                        $False { 'Disabled' }
                                                    }
                                                    'Average Bandwidth (kbit/s)' = $VSSwitch.ExtensionData.Spec.Policy.ShapingPolicy.AverageBandwidth
                                                    'Peak Bandwidth (kbit/s)' = $VSSwitch.ExtensionData.Spec.Policy.ShapingPolicy.PeakBandwidth
                                                    'Burst Size (KB)' = $VSSwitch.ExtensionData.Spec.Policy.ShapingPolicy.BurstSize
                                                }
                                            }
                                            $VssTrafficShapingPolicy | Sort-Object 'Virtual Switch' | Table -Name "$VMHost Virtual Switch Traffic Shaping Policy"
                                        }
                                        #endregion ESXi Host Virtual Switch Traffic Shaping Policy

                                        #region ESXi Host Virtual Switch Teaming & Failover
                                        $VssNicTeamingPolicy = $VSSwitches | Get-NicTeamingPolicy
                                        if ($VssNicTeamingPolicy) {
                                            #region Virtual Switch Teaming & Failover Section
                                            Section -Style Heading5 'Virtual Switch Teaming & Failover' {
                                                $VssNicTeaming = foreach ($VssNicTeam in $VssNicTeamingPolicy) {
                                                    [PSCustomObject]@{
                                                        'Virtual Switch' = $VssNicTeam.VirtualSwitch 
                                                        'Load Balancing' = Switch ($VssNicTeam.LoadBalancingPolicy) {
                                                            'LoadbalanceSrcId' { 'Route based on the originating port ID' }
                                                            'LoadbalanceSrcMac' { 'Route based on source MAC hash' }
                                                            'LoadbalanceIP' { 'Route based on IP hash' }
                                                            'ExplicitFailover' { 'Explicit Failover' }
                                                            default { $VssNicTeam.LoadBalancingPolicy }
                                                        }
                                                        'Network Failure Detection' = Switch ($VssNicTeam.NetworkFailoverDetectionPolicy) {
                                                            'LinkStatus' { 'Link status only' }
                                                            'BeaconProbing' { 'Beacon probing' }
                                                            default { $VssNicTeam.NetworkFailoverDetectionPolicy }
                                                        } 
                                                        'Notify Switches' = Switch ($VssNicTeam.NotifySwitches) {
                                                            $true { 'Yes' }
                                                            $false { 'No' }
                                                        }
                                                        'Failback' = Switch ($VssNicTeam.FailbackEnabled) {
                                                            $true { 'Yes' }
                                                            $false { 'No' }
                                                        }
                                                        'Active NICs' = ($VssNicTeam.ActiveNic | Sort-Object) -join [Environment]::NewLine
                                                        'Standby NICs' = ($VssNicTeam.StandbyNic | Sort-Object) -join [Environment]::NewLine
                                                        'Unused NICs' = ($VssNicTeam.UnusedNic | Sort-Object) -join [Environment]::NewLine
                                                    }
                                                }
                                                $VssNicTeaming | Sort-Object 'Virtual Switch' | Table -Name "$VMHost Virtual Switch Teaming & Failover"
                                            }
                                            #endregion Virtual Switch Teaming & Failover Section
                                        }
                                        #endregion ESXi Host Virtual Switch Teaming & Failover
                                
                                        #region ESXi Host Virtual Switch Port Groups
                                        $VssPortgroups = $VSSwitches | Get-VirtualPortGroup -Standard 
                                        if ($VssPortgroups) {
                                            Section -Style Heading5 'Virtual Switch Port Groups' {
                                                $VssPortgroups = foreach ($VssPortgroup in $VssPortgroups) {
                                                    [PSCustomObject]@{
                                                        'Port Group' = $VssPortgroup.Name
                                                        'VLAN ID' = $VssPortgroup.VLanId 
                                                        'Virtual Switch' = $VssPortgroup.VirtualSwitchName
                                                        '# of VMs' = ($VssPortgroup | Get-VM).Count
                                                    }
                                                }
                                                $VssPortgroups | Sort-Object 'Port Group', 'VLAN ID', 'Virtual Switch' | Table -Name "$VMHost Virtual Switch Port Group Information"
                                            }
                                            #endregion ESXi Host Virtual Switch Port Groups               
                                
                                            #region ESXi Host Virtual Switch Port Group Security Policy
                                            $VssPortgroupSecurity = $VSSwitches | Get-VirtualPortGroup | Get-SecurityPolicy 
                                            if ($VssPortgroupSecurity) {
                                                #region Virtual Port Group Security Policy Section
                                                Section -Style Heading5 'Virtual Switch Port Group Security' {
                                                    $VssPortgroupSecurity = foreach ($VssPortgroupSec in $VssPortgroupSecurity) {
                                                        [PSCustomObject]@{
                                                            'Port Group' = $VssPortgroupSec.VirtualPortGroup
                                                            'Virtual Switch' = $VssPortgroupSec.virtualportgroup.virtualswitchname
                                                            'Promiscuous Mode' = Switch ($VssPortgroupSec.AllowPromiscuous) {
                                                                $true { 'Accept' }
                                                                $false { 'Reject' }
                                                            }
                                                            'MAC Changes' = Switch ($VssPortgroupSec.MacChanges) {
                                                                $true { 'Accept' }
                                                                $false { 'Reject' }
                                                            }
                                                            'Forged Transmits' = Switch ($VssPortgroupSec.ForgedTransmits) {
                                                                $true { 'Accept' }
                                                                $false { 'Reject' }
                                                            } 
                                                        }
                                                    }
                                                    $VssPortgroupSecurity | Sort-Object 'Port Group', 'Virtual Switch' | Table -Name "$VMHost Virtual Switch Port Group Security Policy" 
                                                }
                                                #endregion Virtual Port Group Security Policy Section
                                            }
                                            #endregion ESXi Host Virtual Switch Port Group Security Policy 
                                                                                        
                                            #region ESXi Host Virtual Switch Port Group Traffic Shaping Policy
                                            Section -Style Heading5 'Virtual Switch Port Group Traffic Shaping' {    
                                                $VssPortgroupTrafficShapingPolicy = foreach ($VssPortgroup in $VssPortgroups) {
                                                    [PSCustomObject]@{
                                                        'Port Group' = $VssPortgroup.Name 
                                                        'Virtual Switch' = $VssPortgroup.VirtualSwitchName
                                                        'Status' = Switch ($VssPortgroup.ExtensionData.Spec.Policy.ShapingPolicy.Enabled) {
                                                            $True { 'Enabled' }
                                                            $False { 'Disabled' }
                                                            $null { 'Inherited' }
                                                        }
                                                        'Average Bandwidth (kbit/s)' = $VssPortgroup.ExtensionData.Spec.Policy.ShapingPolicy.AverageBandwidth
                                                        'Peak Bandwidth (kbit/s)' = $VssPortgroup.ExtensionData.Spec.Policy.ShapingPolicy.PeakBandwidth
                                                        'Burst Size (KB)' = $VssPortgroup.ExtensionData.Spec.Policy.ShapingPolicy.BurstSize
                                                    }
                                                }
                                                $VssPortgroupTrafficShapingPolicy | Sort-Object 'Port Group', 'Virtual Switch' | Table -Name "$VMHost Virtual Switch Port Group Traffic Shaping Policy"
                                            }
                                            #endregion ESXi Host Virtual Switch Port Group Traffic Shaping Policy
                                
                                            #region ESXi Host Virtual Switch Port Group Teaming & Failover
                                            $VssPortgroupNicTeaming = $VSSwitches | Get-VirtualPortGroup | Get-NicTeamingPolicy 
                                            if ($VssPortgroupNicTeaming) {
                                                #region Virtual Switch Port Group Teaming & Failover Section
                                                Section -Style Heading5 'Virtual Switch Port Group Teaming & Failover' {
                                                    $VssPortgroupNicTeaming = foreach ($VssPortgroupNicTeam in $VssPortgroupNicTeaming) {
                                                        [PSCustomObject]@{
                                                            'Port Group' = $VssPortgroupNicTeam.VirtualPortGroup
                                                            'Virtual Switch' = $VssPortgroupNicTeam.virtualportgroup.virtualswitchname 
                                                            'Load Balancing' = Switch ($VssPortgroupNicTeam.LoadBalancingPolicy) {
                                                                'LoadbalanceSrcId' { 'Route based on the originating port ID' }
                                                                'LoadbalanceSrcMac' { 'Route based on source MAC hash' }
                                                                'LoadbalanceIP' { 'Route based on IP hash' }
                                                                'ExplicitFailover' { 'Explicit Failover' }
                                                                default { $VssPortgroupNicTeam.LoadBalancingPolicy }
                                                            }
                                                            'Network Failure Detection' = Switch ($VssPortgroupNicTeam.NetworkFailoverDetectionPolicy) {
                                                                'LinkStatus' { 'Link status only' }
                                                                'BeaconProbing' { 'Beacon probing' }
                                                                default { $VssPortgroupNicTeam.NetworkFailoverDetectionPolicy }
                                                            }  
                                                            'Notify Switches' = Switch ($VssPortgroupNicTeam.NotifySwitches) {
                                                                $true { 'Yes' }
                                                                $false { 'No' }
                                                            }
                                                            'Failback' = Switch ($VssPortgroupNicTeam.FailbackEnabled) {
                                                                $true { 'Yes' }
                                                                $false { 'No' }
                                                            } 
                                                            'Active NICs' = ($VssPortgroupNicTeam.ActiveNic | Sort-Object) -join [Environment]::NewLine
                                                            'Standby NICs' = ($VssPortgroupNicTeam.StandbyNic | Sort-Object) -join [Environment]::NewLine
                                                            'Unused NICs' = ($VssPortgroupNicTeam.UnusedNic | Sort-Object) -join [Environment]::NewLine
                                                        }
                                                    }
                                                    $VssPortgroupNicTeaming | Sort-Object 'Port Group', 'Virtual Switch' | Table -Name "$VMHost Virtual Switch Port Group Teaming & Failover"
                                                }
                                                #endregion Virtual Switch Port Group Teaming & Failover Section
                                            }
                                            #endregion ESXi Host Virtual Switch Port Group Teaming & Failover
                                        }
                                    }
                                    #endregion Section Standard Virtual Switches 
                                }
                                #endregion ESXi Host Standard Virtual Switches

                                #region Distributed Virtual Switch Section
                                # Create Distributed Switch Section if they exist
                                $VDSwitches = Get-VDSwitch -Server $ESXi
                                if ($VDSwitches) {
                                    Section -Style Heading3 'Distributed Virtual Switches' {
                                        #region Distributed Virtual Switch Informative Information
                                        if ($InfoLevel.Network -eq 2) {
                                            $VDSInfo = foreach ($VDS in $VDSwitches) {
                                                [PSCustomObject]@{
                                                    'Distributed Switch' = $VDS.Name
                                                    '# of Uplinks' = $VDS.NumUplinkPorts
                                                    '# of Ports' = $VDS.NumPorts 
                                                    '# of Hosts' = $VDS.ExtensionData.Summary.HostMember.Count
                                                    '# of VMs' = $VDS.ExtensionData.Summary.VM.Count
                                                }
                                            }    
                                            $VDSInfo | Table -Name 'Distributed Switch Information'
                                        }    
                                        #endregion Distributed Switch Informative Information

                                        #region Distributed Switch Detailed Information
                                        if ($InfoLevel.Network -ge 3) {
                                            ## TODO: LACP, NetFlow, NIOC
                                            foreach ($VDS in ($VDSwitches)) {
                                                #region VDS Section
                                                Section -Style Heading4 $VDS {
                                                    #region Distributed Switch General Properties  
                                                    $VDSwitchDetail = [PSCustomObject]@{
                                                        'Distributed Switch' = $VDS.Name
                                                        'ID' = $VDS.Id
                                                        'Number of Ports' = $VDS.NumPorts
                                                        'Number of Port Groups' = $VDS.ExtensionData.Summary.PortGroupName.Count 
                                                        'Number of VMs' = $VDS.ExtensionData.Summary.VM.Count 
                                                        'MTU' = $VDS.Mtu
                                                        'Network I/O Control' = Switch ($VDS.ExtensionData.Config.NetworkResourceManagementEnabled) {
                                                            $true { 'Enabled' }
                                                            $false { 'Disabled' }
                                                        } 
                                                        'Discovery Protocol' = $VDS.LinkDiscoveryProtocol
                                                        'Discovery Protocol Operation' = $VDS.LinkDiscoveryProtocolOperation
                                                    }

                                                    #region Network Advanced Detail Information
                                                    if ($InfoLevel.Network -ge 4) {
                                                        $VDSwitchDetail | ForEach-Object {
                                                            $VDSwitchVMs = $VDS | Get-VM | Sort-Object 
                                                            Add-Member -InputObject $_ -MemberType NoteProperty -Name 'Virtual Machines' -Value ($VDSwitchVMs.Name -join ', ')
                                                        }
                                                    }
                                                    #endregion Network Advanced Detail Information
                                                    $VDSwitchDetail | Table -Name "$VDS Distributed Switch General Properties" -List -ColumnWidths 50, 50 
                                                    #endregion Distributed Switch General Properties

                                                    #region Distributed Switch Uplink Ports
                                                    $VdsUplinks = $VDS | Get-VDPortgroup | Where-Object { $_.IsUplink -eq $true } | Get-VDPort
                                                    if ($VdsUplinks) {
                                                        Section -Style Heading4 'Distributed Switch Uplink Ports' {
                                                            $VdsUplinkDetail = foreach ($VdsUplink in $VdsUplinks) {
                                                                [PSCustomObject]@{
                                                                    'Distributed Switch' = $VdsUplink.Switch
                                                                    'Uplink Name' = $VdsUplink.Name
                                                                    'Physical Network Adapter' = $VdsUplink.ConnectedEntity
                                                                    'Uplink Port Group' = $VdsUplink.Portgroup
                                                                }
                                                            }
                                                            $VdsUplinkDetail | Sort-Object 'Distributed Switch', 'Uplink Name' | Table -Name "$VDS Distributed Switch Uplink Ports"
                                                        }
                                                    }
                                                    #endregion Distributed Virtual Switch Uplink Ports
                                                    
                                                    #region Distributed Switch Port Groups
                                                    $VDSPortgroups = $VDS | Get-VDPortgroup
                                                    if ($VDSPortgroups) {
                                                        Section -Style Heading4 'Distributed Switch Port Groups' {
                                                            $VDSPortgroupDetail = foreach ($VDSPortgroup in $VDSPortgroups) {
                                                                [PSCustomObject]@{
                                                                    'Port Group' = $VDSPortgroup.Name
                                                                    'Distributed Switch' = $VDSPortgroup.VDSwitch.Name
                                                                    'VLAN Configuration' = Switch ($VDSPortgroup.VlanConfiguration) {
                                                                        $null { '--' }
                                                                        default { $VDSPortgroup.VlanConfiguration }
                                                                    }
                                                                    'Port Binding' = $VDSPortgroup.PortBinding
                                                                }
                                                            }
                                                            $VDSPortgroupDetail | Sort-Object 'Port Group' | Table -Name "$VDS Distributed Switch Port Groups" 
                                                        }
                                                    }
                                                    #endregion Distributed Switch Port Groups

                                                    #region Distributed Switch Private VLANs
                                                    $VDSwitchPrivateVLANs = $VDS | Get-VDSwitchPrivateVlan
                                                    if ($VDSwitchPrivateVLANs) {
                                                        Section -Style Heading4 'Distributed Switch Private VLANs' {
                                                            $VDSPvlan = foreach ($VDSwitchPrivateVLAN in $VDSwitchPrivateVLANs) {
                                                                [PSCustomObject]@{
                                                                    'Primary VLAN ID' = $VDSwitchPrivateVLAN.PrimaryVlanId
                                                                    'Private VLAN Type' = $VDSwitchPrivateVLAN.PrivateVlanType
                                                                    'Secondary VLAN ID' = $VDSwitchPrivateVLAN.SecondaryVlanId
                                                                }
                                                            }
                                                            $VDSPvlan | Sort-Object 'Primary VLAN ID', 'Secondary VLAN ID' | Table -Name "$VDS Distributed Switch Private VLANs"
                                                        }
                                                    }
                                                    #endregion Distributed Switch Private VLANs  
                                                }
                                                #endregion VDS Section
                                            }
                                        }
                                        #endregion Distributed Virtual Switch Detailed Information
                                    }
                                }
                                #endregion Distributed Virtual Switch Section
                            }                
                            #endregion ESXi Host Network Section

                            #region ESXi Host Security Section
                            Section -Style Heading2 'Security' {
                                Paragraph "The following section details the host security configuration for $($VMHost.ExtensionData.Name)."
                                #region ESXi Host Lockdown Mode
                                if ($VMHost.ExtensionData.Config.LockdownMode -ne $null) {
                                    Section -Style Heading3 'Lockdown Mode' {
                                        $LockdownMode = [PSCustomObject]@{
                                            'Lockdown Mode' = Switch ($VMHost.ExtensionData.Config.LockdownMode) {
                                                'lockdownDisabled' { 'Disabled' }
                                                'lockdownNormal' { 'Enabled (Normal)' }
                                                'lockdownStrict' { 'Enabled (Strict)' }
                                                default { $VMHost.ExtensionData.Config.LockdownMode }
                                            }
                                        }
                                        if ($Healthcheck.VMHost.LockdownMode) {
                                            $LockdownMode | Where-Object { $_.'Lockdown Mode' -eq 'Disabled' } | Set-Style -Style Warning -Property 'Lockdown Mode'
                                        }
                                        $LockdownMode | Table -Name "$($VMHost.ExtensionData.Name) Lockdown Mode" -List -ColumnWidths 50, 50
                                    }
                                }
                                #endregion ESXi Host Lockdown Mode

                                #region ESXi Host Services
                                Section -Style Heading3 'Services' {
                                    $VMHostServices = $VMHost | Get-VMHostService
                                    $Services = foreach ($VMHostService in $VMHostServices) {
                                        [PSCustomObject]@{
                                            'Service' = $VMHostService.Label
                                            'Daemon' = Switch ($VMHostService.Running) {
                                                $true { 'Running' }
                                                $false { 'Stopped' }
                                            }
                                            'Startup Policy' = Switch ($VMHostService.Policy) {
                                                'automatic' { 'Start and stop with port usage' }
                                                'on' { 'Start and stop with host' }
                                                'off' { 'Start and stop manually' }
                                                default { $VMHostService.Policy }
                                            }
                                        }
                                    }
                                    if ($Healthcheck.VMHost.NTP) {
                                        $Services | Where-Object { ($_.'Service' -eq 'NTP Daemon') -and ($_.Daemon -eq 'Stopped') } | Set-Style -Style Critical -Property 'Daemon'
                                        $Services | Where-Object { ($_.'Service' -eq 'NTP Daemon') -and ($_.'Startup Policy' -ne 'Start and stop with host') } | Set-Style -Style Critical -Property 'Startup Policy'
                                    }
                                    if ($Healthcheck.VMHost.SSH) {
                                        $Services | Where-Object { ($_.'Service' -eq 'SSH') -and ($_.Daemon -eq 'Running') } | Set-Style -Style Warning -Property 'Daemon'
                                        $Services | Where-Object { ($_.'Service' -eq 'SSH') -and ($_.'Startup Policy' -ne 'Start and stop manually') } | Set-Style -Style Warning -Property 'Startup Policy'
                                    }
                                    if ($Healthcheck.VMHost.ESXiShell) {
                                        $Services | Where-Object { ($_.'Service' -eq 'ESXi Shell') -and ($_.Daemon -eq 'Running') } | Set-Style -Style Warning -Property 'Daemon'
                                        $Services | Where-Object { ($_.'Service' -eq 'ESXi Shell') -and ($_.'Startup Policy' -ne 'Start and stop manually') } | Set-Style -Style Warning -Property 'Startup Policy'
                                    }
                                    $Services | Sort-Object 'Service' | Table -Name "$($VMHost.ExtensionData.Name) Services" 
                                }
                                #endregion ESXi Host Services

                                #region ESXi Host Advanced Detail Information
                                if ($InfoLevel.VMHost -ge 4) {
                                    #region ESXi Host Firewall
                                    $VMHostFirewallExceptions = $VMHost | Get-VMHostFirewallException
                                    if ($VMHostFirewallExceptions) {
                                        #region Friewall Section
                                        Section -Style Heading3 'Firewall' {
                                            $VMHostFirewall = foreach ($VMHostFirewallException in $VMHostFirewallExceptions) {
                                                [PScustomObject]@{
                                                    'Service' = $VMHostFirewallException.Name
                                                    'Status' = Switch ($VMHostFirewallException.Enabled) {
                                                        $true { 'Enabled' }
                                                        $false { 'Disabled' }
                                                    }
                                                    'Incoming Ports' = $VMHostFirewallException.IncomingPorts
                                                    'Outgoing Ports' = $VMHostFirewallException.OutgoingPorts
                                                    'Protocols' = $VMHostFirewallException.Protocols
                                                    'Daemon' = Switch ($VMHostFirewallException.ServiceRunning) {
                                                        $true { 'Running' }
                                                        $false { 'Stopped' }
                                                        $null { 'N/A' }
                                                        default { $VMHostFirewallException.ServiceRunning }
                                                    }
                                                }
                                            }
                                            $VMHostFirewall | Sort-Object 'Service' | Table -Name "$($VMHost.ExtensionData.Name) Firewall Configuration" 
                                        }
                                        #endregion Friewall Section
                                    }
                                    #endregion ESXi Host Firewall
    
                                    #region ESXi Host Authentication
                                    $AuthServices = $VMHost | Get-VMHostAuthentication
                                    if ($AuthServices.DomainMembershipStatus) {
                                        Section -Style Heading3 'Authentication Services' {
                                            $AuthServices = $AuthServices | Select-Object Domain, @{L = 'Domain Membership'; E = { $_.DomainMembershipStatus } }, @{L = 'Trusted Domains'; E = { $_.TrustedDomains } }
                                            $AuthServices | Table -Name "$($VMHost.ExtensionData.Name) Authentication Services" -ColumnWidths 25, 25, 50 
                                        }    
                                    }
                                    #endregion ESXi Host Authentication
                                }
                                #endregion ESXi Host Advanced Detail Information
                            }
                            #endregion ESXi Host Security Section

                            #region Virtual Machine Section
                            if ($InfoLevel.VM -ge 1) {
                                if ($VMs) {
                                    Section -Style Heading2 'Virtual Machines' {
                                        Paragraph "The following section details the configuration of virtual machines managed by $($VMHost.ExtensionData.Name)."
                                        #region Virtual Machine Summary Information
                                        if ($InfoLevel.VM -eq 1) {
                                            BlankLine
                                            $VMSummary = [PSCustomObject]@{
                                                'Total VMs' = $VMs.Count
                                                'Total vCPUs' = ($VMs | Measure-Object -Property NumCpu -Sum).Sum
                                                'Total Memory' = "$([math]::Round(($VMs | Measure-Object -Property MemoryGB -Sum).Sum, 2)) GB"
                                                'Total Provisioned Space' = "$([math]::Round(($VMs | Measure-Object -Property ProvisionedSpaceGB -Sum).Sum, 2)) GB"
                                                'Total Used Space' = "$([math]::Round(($VMs | Measure-Object -Property UsedSpaceGB -Sum).Sum, 2)) GB"
                                                'VMs Powered On' = ($VMs | Where-Object { $_.PowerState -eq 'PoweredOn' }).Count
                                                'VMs Powered Off' = ($VMs | Where-Object { $_.PowerState -eq 'PoweredOff' }).Count
                                                'VMs Suspended' = ($VMs | Where-Object { $_.PowerState -eq 'Suspended' }).Count
                                                'VMs with Snapshots' = ($VMs | Where-Object { $_.ExtensionData.Snapshot }).Count
                                                'Guest Operating System Types' = (($VMs | Get-View).Summary.Config.GuestFullName | Select-Object -Unique).Count
                                                'VM Tools OK' = ($VMs | Where-Object { $_.ExtensionData.Guest.ToolsStatus -eq 'toolsOK' }).Count
                                                'VM Tools Old' = ($VMs | Where-Object { $_.ExtensionData.Guest.ToolsStatus -eq 'toolsOld' }).Count
                                                'VM Tools Not Running' = ($VMs | Where-Object { $_.ExtensionData.Guest.ToolsStatus -eq 'toolsNotRunning' }).Count
                                                'VM Tools Not Installed' = ($VMs | Where-Object { $_.ExtensionData.Guest.ToolsStatus -eq 'toolsNotInstalled' }).Count
                                            }
                                            $VMSummary | Table -List -Name 'VM Summary' -ColumnWidths 50, 50
                                        }
                                        #endregion Virtual Machine Summary Information

                                        #region Virtual Machine Informative Information
                                        if ($InfoLevel.VM -eq 2) {
                                            BlankLine
                                            $VMSnapshotList = $VMs.Extensiondata.Snapshot.RootSnapshotList
                                            $VMInfo = foreach ($VM in $VMs) {
                                                $VMView = $VM | Get-View
                                                [PSCustomObject]@{
                                                    'Virtual Machine' = $VM.Name
                                                    'Power State' = Switch ($VM.PowerState) {
                                                        'PoweredOn' { 'On' }
                                                        'PoweredOff' { 'Off' }
                                                        default { $VM.PowerState }
                                                    }
                                                    'IP Address' = Switch ($VMView.Guest.IpAddress) {
                                                        $null { '--' }
                                                        default { $VMView.Guest.IpAddress }
                                                    }
                                                    'vCPUs' = $VM.NumCpu
                                                    'Memory GB' = [math]::Round(($VM.MemoryGB), 0)
                                                    'Provisioned GB' = [math]::Round(($VM.ProvisionedSpaceGB), 0)
                                                    'Used GB' = [math]::Round(($VM.UsedSpaceGB), 0)
                                                    'HW Version' = ($VM.HardwareVersion).Replace('vmx-', 'v')
                                                    'VM Tools Status' = Switch ($VMView.Guest.ToolsStatus) {
                                                        'toolsOld' { 'Old' }
                                                        'toolsOK' { 'OK' }
                                                        'toolsNotRunning' { 'Not Running' }
                                                        'toolsNotInstalled' { 'Not Installed' }
                                                        default { $VMView.Guest.ToolsStatus }
                                                    }         
                                                }
                                            }
                                            if ($Healthcheck.VM.VMToolsStatus) {
                                                $VMInfo | Where-Object { $_.'VM Tools Status' -ne 'OK' } | Set-Style -Style Warning -Property 'VM Tools Status'
                                            }
                                            if ($Healthcheck.VM.PowerState) {
                                                $VMInfo | Where-Object { $_.'Power State' -ne 'On' } | Set-Style -Style Warning -Property 'Power State'
                                            }
                                            $VMInfo | Table -Name 'VM Informative Information'

                                            #region VM Snapshot Information
                                            if ($VMSnapshotList -and $Options.ShowVMSnapshots) {
                                                Section -Style Heading3 'Snapshots' {
                                                    $VMSnapshotInfo = foreach ($VMSnapshot in $VMSnapshotList) {
                                                        [PSCustomObject]@{
                                                            'Virtual Machine' = $VMLookup."$($VMSnapshot.VM)"
                                                            'Snapshot Name' = $VMSnapshot.Name
                                                            'Description' = $VMSnapshot.Description
                                                            'Days Old' = ((Get-Date).ToUniversalTime() - $VMSnapshot.CreateTime).Days
                                                        } 
                                                    }
                                                    if ($Healthcheck.VM.VMSnapshots) {
                                                        $VMSnapshotInfo | Where-Object { $_.'Days Old' -ge 7 } | Set-Style -Style Warning 
                                                        $VMSnapshotInfo | Where-Object { $_.'Days Old' -ge 14 } | Set-Style -Style Critical
                                                    }
                                                    $VMSnapshotInfo | Table -Name 'VM Snapshot Information'
                                                }
                                            }
                                            #endregion VM Snapshot Information
                                        }
                                        #endregion Virtual Machine Informative Information

                                        #region Virtual Machine Detailed Information
                                        if ($InfoLevel.VM -ge 3) {
                                            foreach ($VM in $VMs) {
                                                Section -Style Heading3 $VM.name {
                                                    $VMUptime = @()
                                                    $VMUptime = Get-Uptime -VM $VM
                                                    $VMSpbmPolicy = $VMSpbmConfig | Where-Object { $_.entity -eq $vm }
                                                    $VMView = $VM | Get-View
                                                    $VMSnapshotList = $vmview.Snapshot.RootSnapshotList
                                                    $VMDetail = [PSCustomObject]@{
                                                        'Virtual Machine' = $VM.Name
                                                        'ID' = $VM.Id 
                                                        'Operating System' = $VMView.Summary.Config.GuestFullName
                                                        'Hardware Version' = ($VM.HardwareVersion).Replace('vmx-', 'v')
                                                        'Power State' = Switch ($VM.PowerState) {
                                                            'PoweredOn' { 'On' }
                                                            'PoweredOff' { 'Off' }
                                                            default { $VM.PowerState }
                                                        }
                                                        'Connection State' = $TextInfo.ToTitleCase($VM.ExtensionData.Runtime.ConnectionState)
                                                        'VM Tools Status' = Switch ($VMView.Guest.ToolsStatus) {
                                                            'toolsOld' { 'Old' }
                                                            'toolsOK' { 'OK' }
                                                            'toolsNotRunning' { 'Not Running' }
                                                            'toolsNotInstalled' { 'Not Installed' }
                                                            default { $VMView.Guest.ToolsStatus }
                                                        }
                                                        'Fault Tolerance State' = Switch ($VMView.Runtime.FaultToleranceState) {
                                                            'notConfigured' { 'Not Configured' }
                                                            'needsSecondary' { 'Needs Secondary' }
                                                            'running' { 'Running' }
                                                            'disabled' { 'Disabled' }
                                                            'starting' { 'Starting' }
                                                            'enabled' { 'Enabled' }
                                                            default { $VMview.Runtime.FaultToleranceState }
                                                        } 
                                                        'vCPUs' = $VM.NumCpu
                                                        'Cores per Socket' = $VM.CoresPerSocket
                                                        'CPU Shares' = "$($VM.VMResourceConfiguration.CpuSharesLevel) / $($VM.VMResourceConfiguration.NumCpuShares)"
                                                        'CPU Reservation' = $VM.VMResourceConfiguration.CpuReservationMhz
                                                        'CPU Limit' = "$($VM.VMResourceConfiguration.CpuReservationMhz) MHz" 
                                                        'CPU Hot Add' = Switch ($VMView.Config.CpuHotAddEnabled) {
                                                            $true { 'Enabled' }
                                                            $false { 'Disabled' }
                                                        }
                                                        'CPU Hot Remove' = Switch ($VMView.Config.CpuHotRemoveEnabled) {
                                                            $true { 'Enabled' }
                                                            $false { 'Disabled' }
                                                        } 
                                                        'Memory Allocation' = "$([math]::Round(($VM.memoryGB), 2)) GB" 
                                                        'Memory Shares' = "$($VM.VMResourceConfiguration.MemSharesLevel) / $($VM.VMResourceConfiguration.NumMemShares)"
                                                        'Memory Hot Add' = Switch ($VMView.Config.MemoryHotAddEnabled) {
                                                            $true { 'Enabled' }
                                                            $false { 'Disabled' }
                                                        }
                                                        'vNICs' = $VMView.Summary.Config.NumEthernetCards
                                                        'DNS Name' = if ($VMView.Guest.HostName) {
                                                            $VMView.Guest.HostName
                                                        } else {
                                                            '--'
                                                        }
                                                        'Networks' = if ($VMView.Guest.Net.Network) {
                                                            (($VMView.Guest.Net | Where-Object { $_.Network -ne $null } | Select-Object Network | Sort-Object Network).Network -join ', ')
                                                        } else {
                                                            '--'
                                                        }
                                                        'IP Address' = if ($VMView.Guest.Net.IpAddress) {
                                                            (($VMView.Guest.Net | Where-Object { ($_.Network -ne $null) -and ($_.IpAddress -ne $null) } | Select-Object IpAddress | Sort-Object IpAddress).IpAddress -join ', ')
                                                        } else {
                                                            '--'
                                                        }
                                                        'MAC Address' = if ($VMView.Guest.Net.MacAddress) {
                                                            (($VMView.Guest.Net | Where-Object { $_.Network -ne $null } | Select-Object -Property MacAddress).MacAddress -join ', ')
                                                        } else {
                                                            '--'
                                                        }
                                                        'vDisks' = $VMView.Summary.Config.NumVirtualDisks
                                                        'Provisioned Space' = "$([math]::Round(($VM.ProvisionedSpaceGB), 2)) GB"
                                                        'Used Space' = "$([math]::Round(($VM.UsedSpaceGB), 2)) GB"
                                                        'Changed Block Tracking' = Switch ($VMView.Config.ChangeTrackingEnabled) {
                                                            $true { 'Enabled' }
                                                            $false { 'Disabled' }
                                                        }
                                                    }
                                                    $MemberProps = @{
                                                        'InputObject' = $VMDetail
                                                        'MemberType' = 'NoteProperty'
                                                    }
                                                    #if ($VMView.Config.CreateDate) {
                                                    #    Add-Member @MemberProps -Name 'Creation Date' -Value ($VMView.Config.CreateDate).ToLocalTime()
                                                    #}
                                                    if ($VM.Notes) {
                                                        Add-Member @MemberProps -Name 'Notes' -Value $VM.Notes  
                                                    }
                                                    if ($VMView.Runtime.BootTime) {
                                                        Add-Member @MemberProps -Name 'Boot Time' -Value ($VMView.Runtime.BootTime).ToLocalTime()
                                                    }
                                                    if ($VMUptime.UptimeDays) {
                                                        Add-Member @MemberProps -Name 'Uptime Days' -Value $VMUptime.UptimeDays
                                                    }

                                                    #region VM Health Checks
                                                    if ($Healthcheck.VM.VMToolsStatus) {
                                                        $VMDetail | Where-Object { $_.'VM Tools Status' -ne 'OK' } | Set-Style -Style Warning -Property 'VM Tools Status'
                                                    }
                                                    if ($Healthcheck.VM.PowerState) {
                                                        $VMDetail | Where-Object { $_.'Power State' -ne 'On' } | Set-Style -Style Warning -Property 'Power State'
                                                    }
                                                    if ($Healthcheck.VM.ConnectionState) {
                                                        $VMDetail | Where-Object { $_.'Connection State' -ne 'Connected' } | Set-Style -Style Critical -Property 'Connection State'
                                                    }
                                                    if ($Healthcheck.VM.CpuHotAdd) {
                                                        $VMDetail | Where-Object { $_.'CPU Hot Add' -eq 'Enabled' } | Set-Style -Style Warning -Property 'CPU Hot Add'
                                                    }
                                                    if ($Healthcheck.VM.CpuHotRemove) {
                                                        $VMDetail | Where-Object { $_.'CPU Hot Remove' -eq 'Enabled' } | Set-Style -Style Warning -Property 'CPU Hot Remove'
                                                    } 
                                                    if ($Healthcheck.VM.MemoryHotAdd) {
                                                        $VMDetail | Where-Object { $_.'Memory Hot Add' -eq 'Enabled' } | Set-Style -Style Warning -Property 'Memory Hot Add'
                                                    } 
                                                    if ($Healthcheck.VM.ChangeBlockTracking) {
                                                        $VMDetail | Where-Object { $_.'Changed Block Tracking' -eq 'Disabled' } | Set-Style -Style Warning -Property 'Changed Block Tracking'
                                                    } 
                                                    #endregion VM Health Checks

                                                    $VMDetail | Table -Name "$($VM.Name) Detailed Information" -List -ColumnWidths 50, 50

                                                    if ($InfoLevel.VM -ge 4) {
                                                        $VMnics = $VM.Guest.Nics | Where-Object { $_.Device -ne $null } | Sort-Object Device
                                                        $VMHdds = $VMHardDisks | Where-Object { $_.ParentId -eq $VM.Id } | Sort-Object Name
                                                        $SCSIControllers = $VMView.Config.Hardware.Device | Where-Object { $_.DeviceInfo.Label -match "SCSI Controller" }
                                                        $VMGuestVols = $VM.Guest.Disks | Sort-Object Path
                                                        if ($VMnics) {
                                                            Section -Style Heading4 "Network Adapters" {
                                                                $VMnicInfo = foreach ($VMnic in $VMnics) {
                                                                    [PSCustomObject]@{
                                                                        'Adapter' = $VMnic.Device
                                                                        'Connected' = $VMnic.Connected
                                                                        'Network Name' = Switch -wildcard ($VMnic.Device.NetworkName) {
                                                                            'dvportgroup*' { $VDPortgroupLookup."$($VMnic.Device.NetworkName)" }
                                                                            default { $VMnic.Device.NetworkName }
                                                                        }
                                                                        'Adapter Type' = $VMnic.Device.Type
                                                                        'IP Address' = $VMnic.IpAddress -join [Environment]::NewLine
                                                                        'MAC Address' = $VMnic.Device.MacAddress
                                                                    }
                                                                }
                                                                $VMnicInfo | Table -Name "$($VM.Name) Network Adapters"
                                                            }
                                                        }
                                                        if ($SCSIControllers) {
                                                            Section -Style Heading4 "SCSI Controllers" {
                                                                $VMScsiControllers = foreach ($VMSCSIController in $SCSIControllers) {
                                                                    [PSCustomObject]@{
                                                                        'Device' = $VMSCSIController.DeviceInfo.Label
                                                                        'Controller Type' = $VMSCSIController.DeviceInfo.Summary
                                                                        'Bus Sharing' = Switch ($VMSCSIController.SharedBus) {
                                                                            'noSharing' { 'None' }
                                                                            default { $VMSCSIController.SharedBus }
                                                                        }
                                                                    }
                                                                }
                                                                $VMScsiControllers | Sort-Object 'Device' | Table -Name "$($VM.Name) SCSI Controllers"
                                                            }
                                                        }
                                                        if ($VMHdds) {
                                                            Section -Style Heading4 "Hard Disks" {
                                                                If ($InfoLevel.VM -eq 4) {
                                                                    $VMHardDiskInfo = foreach ($VMHdd in $VMHdds) {
                                                                        $SCSIDevice = $VMView.Config.Hardware.Device | Where-Object { $_.Key -eq $VMHdd.ExtensionData.Key -and $_.Backing.FileName -eq $VMHdd.FileName }
                                                                        $SCSIController = $SCSIControllers | Where-Object { $SCSIDevice.ControllerKey -eq $_.Key }
                                                                        [PSCustomObject]@{
                                                                            'Disk' = $VMHdd.Name
                                                                            'Datastore' = $VMHdd.FileName.Substring($VMHdd.Filename.IndexOf("[") + 1, $VMHdd.Filename.IndexOf("]") - 1)
                                                                            'Capacity' = "$([math]::Round(($VMHdd.CapacityGB), 2)) GB"
                                                                            'Disk Provisioning' = Switch ($VMHdd.StorageFormat) {
                                                                                'EagerZeroedThick' { 'Thick Eager Zeroed' }
                                                                                'LazyZeroedThick' { 'Thick Lazy Zeroed' }
                                                                                $null { '--' }
                                                                                default { $VMHdd.StorageFormat }
                                                                            }
                                                                            'Disk Type' = Switch ($VMHdd.DiskType) {
                                                                                'RawPhysical' { 'Physical RDM' }
                                                                                'RawVirtual' { "Virtual RDM" }
                                                                                'Flat' { 'VMDK' }
                                                                                default { $VMHdd.DiskType }
                                                                            }
                                                                            'Disk Mode' = Switch ($VMHdd.Persistence) {
                                                                                'IndependentPersistent' { 'Independent - Persistent' }
                                                                                'IndependentNonPersistent' { 'Independent - Nonpersistent' }
                                                                                'Persistent' { 'Dependent' }
                                                                                default { $VMHdd.Persistence }
                                                                            }
                                                                        }
                                                                    }
                                                                    $VMHardDiskInfo | Table -Name "$($VM.Name) Hard Disk Information"
                                                                } else {
                                                                    foreach ($VMHdd in $VMHdds) {
                                                                        Section -Style Heading4 "$($VMHdd.Name)" {
                                                                            $SCSIDevice = $VMView.Config.Hardware.Device | Where-Object { $_.Key -eq $VMHdd.ExtensionData.Key -and $_.Backing.FileName -eq $VMHdd.FileName }
                                                                            $SCSIController = $SCSIControllers | Where-Object { $SCSIDevice.ControllerKey -eq $_.Key }
                                                                            $VMHardDiskInfo = [PSCustomObject]@{
                                                                                'Datastore' = $VMHdd.FileName.Substring($VMHdd.Filename.IndexOf("[") + 1, $VMHdd.Filename.IndexOf("]") - 1)
                                                                                'Capacity' = "$([math]::Round(($VMHdd.CapacityGB), 2)) GB"
                                                                                'Disk Path' = $VMHdd.Filename.Substring($VMHdd.Filename.IndexOf("]") + 2)
                                                                                'Disk Shares' = "$($TextInfo.ToTitleCase($VMHdd.ExtensionData.Shares.Level)) / $($VMHdd.ExtensionData.Shares.Shares)"
                                                                                'Disk Limit IOPs' = Switch ($VMHdd.ExtensionData.StorageIOAllocation.Limit) {
                                                                                    '-1' { 'Unlimited' }
                                                                                    default { $VMHdd.ExtensionData.StorageIOAllocation.Limit }
                                                                                }
                                                                                'Disk Provisioning' = Switch ($VMHdd.StorageFormat) {
                                                                                    'EagerZeroedThick' { 'Thick Eager Zeroed' }
                                                                                    'LazyZeroedThick' { 'Thick Lazy Zeroed' }
                                                                                    $null { '--' }
                                                                                    default { $VMHdd.StorageFormat }
                                                                                }
                                                                                'Disk Type' = Switch ($VMHdd.DiskType) {
                                                                                    'RawPhysical' { 'Physical RDM' }
                                                                                    'RawVirtual' { "Virtual RDM" }
                                                                                    'Flat' { 'VMDK' }
                                                                                    default { $VMHdd.DiskType }
                                                                                }
                                                                                'Disk Mode' = Switch ($VMHdd.Persistence) {
                                                                                    'IndependentPersistent' { 'Independent - Persistent' }
                                                                                    'IndependentNonPersistent' { 'Independent - Nonpersistent' }
                                                                                    'Persistent' { 'Dependent' }
                                                                                    default { $VMHdd.Persistence }
                                                                                }
                                                                                'SCSI Controller' = $SCSIController.DeviceInfo.Label
                                                                                'SCSI Address' = "$($SCSIController.BusNumber):$($VMHdd.ExtensionData.UnitNumber)"
                                                                            }
                                                                            $VMHardDiskInfo | Table -List "$($VM.Name) $($VMHdd.Name) Information" -ColumnWidths 25, 75
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        if ($VMGuestVols) {
                                                            Section -Style Heading4 "Guest Volumes" {
                                                                $VMGuestDiskInfo = foreach ($VMGuestVol in $VMGuestVols) {
                                                                    [PSCustomObject]@{
                                                                        'Path' = $VMGuestVol.Path
                                                                        'Capacity' = "$([math]::Round(($VMGuestVol.CapacityGB), 2)) GB"
                                                                        'Used Space' = "$([math]::Round((($VMGuestVol.CapacityGB) - ($VMGuestVol.FreeSpaceGB)), 2)) GB"
                                                                        'Free Space' = "$([math]::Round($VMGuestVol.FreeSpaceGB, 2)) GB"
                                                                    }
                                                                }
                                                                $VMGuestDiskInfo | Table -Name "$($VM.Name) Guest Volumes" -ColumnWidths 25, 25, 25, 25
                                                            }
                                                        }
                                                    }

                                    
                                                    if ($VMSnapshotList -and $Options.ShowVMSnapshots) {
                                                        Section -Style Heading4 "Snapshots" {
                                                            $VMSnapshots = foreach ($VMSnapshot in $VMSnapshotList) {
                                                                [PSCustomObject]@{
                                                                    'Snapshot Name' = $VMSnapshot.Name
                                                                    'Description' = $VMSnapshot.Description
                                                                    'Days Old' = ((Get-Date).ToUniversalTime() - $VMSnapshot.CreateTime).Days
                                                                } 
                                                            }
                                                            if ($Healthcheck.VM.VMSnapshots) {
                                                                $VMSnapshots | Where-Object { $_.'Days Old' -ge 7 } | Set-Style -Style Warning 
                                                                $VMSnapshots | Where-Object { $_.'Days Old' -ge 14 } | Set-Style -Style Critical
                                                            }
                                                            $VMSnapshots | Table -Name "$($VM.Name) Snapshots"
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        #endregion Virtual Machine Detailed Information
                                    }
                                }
                            }
                            #endregion Virtual Machine Section

                            #region ESXi Host VM Startup/Shutdown Information
                            $VMStartPolicy = $VMHost | Get-VMStartPolicy | Where-Object { $_.StartAction -ne 'None' }
                            if ($VMStartPolicy) {
                                #region VM Startup/Shutdown Section
                                Section -Style Heading2 'VM Startup/Shutdown' {
                                    Paragraph "The following section details the VM startup/shutdown configuration for $($VMHost.ExtensionData.Name)."
                                    BlankLine
                                    $VMStartPolicies = foreach ($VMStartPol in $VMStartPolicy) {
                                        [PSCustomObject]@{
                                            'Start Order' = $VMStartPol.StartOrder
                                            'VM Name' = $VMStartPol.VirtualMachineName
                                            'Startup' = Switch ($VMStartPol.StartAction) {
                                                'PowerOn' { 'Enabled' }
                                                'None' { 'Disabled' }
                                                default { $VMStartPol.StartAction }
                                            }
                                            'Startup Delay' = "$($VMStartPol.StartDelay) seconds"
                                            'VMware Tools' = Switch ($VMStartPol.WaitForHeartbeat) {
                                                $true { 'Continue if VMware Tools is started' }
                                                $false { 'Wait for startup delay' }
                                            }
                                            'Shutdown Behavior' = Switch ($VMStartPol.StopAction) {
                                                'PowerOff' { 'Power Off' }
                                                'GuestShutdown' { 'Guest Shutdown' }
                                                default { $VMStartPol.StopAction }
                                            }
                                            'Shutdown Delay' = "$($VMStartPol.StopDelay) seconds"
                                        }
                                    }
                                    $VMStartPolicies | Table -Name "$($VMHost.ExtensionData.Name) VM Startup/Shutdown Policy" 
                                }
                                #endregion VM Startup/Shutdown Section
                            }
                            #endregion ESXi Host VM Startup/Shutdown Information
                        }
                        #endregion ESXi Host Detailed Information
                    }
                    #endregion Hosts Section
                } # end if ($VMHosts)
            } # end if ($InfoLevel.VMHost -ge 1)          
        } # end if ($ESXi)

        # Disconnect ESXi Server
        $Null = Disconnect-VIServer -Server $ESXi -Confirm:$false -ErrorAction SilentlyContinue

        #region Variable cleanup
        Clear-Variable -Name ESXi
        #endregion Variable cleanup

    } # end foreach ($VIServer in $Target)
}