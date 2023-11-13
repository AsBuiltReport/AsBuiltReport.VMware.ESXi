function Invoke-AsBuiltReport.VMware.ESXi {
    <#
    .SYNOPSIS
        PowerShell script to document the configuration of VMware ESXi servers in Word/HTML/XML/Text formats
    .DESCRIPTION
        Documents the configuration of VMware ESXi servers in Word/HTML/XML/Text formats using PScribo.
    .NOTES
        Version:        1.1.3
        Author:         Tim Carman
        Twitter:        @tpcarman
        Github:         tpcarman
        Credits:        Iain Brighton (@iainbrighton) - PScribo module
    .LINK
        https://github.com/AsBuiltReport/AsBuiltReport.VMware.ESXi
    #>

    param (
        [String[]] $Target,
        [PSCredential] $Credential
    )

    # Check if the required version of VMware PowerCLI is installed
    Get-RequiredModule -Name 'VMware.PowerCLI' -Version '12.3'

    # Import Report Configuration
    $Report = $ReportConfig.Report
    $InfoLevel = $ReportConfig.InfoLevel
    $Options = $ReportConfig.Options
    # Used to set values to TitleCase where required
    $TextInfo = (Get-Culture).TextInfo

    #region Script Body
    #---------------------------------------------------------------------------------------------#
    #                                         SCRIPT BODY                                         #
    #---------------------------------------------------------------------------------------------#
    # Connect to ESXi Server using supplied credentials
    foreach ($VIServer in $Target) {
        try {
            Write-PScriboMessage "Connecting to ESXi Server '$VIServer'."
            $ESXi = Connect-VIServer $VIServer -Credential $Credential -Protocol https -ErrorAction Stop
        } catch {
            Write-Error $_
        }
        #region Generate ESXi report
        if ($ESXi) {
            # Create a lookup hashtable to quickly link VM MoRefs to Names
            # Exclude VMware Site Recovery Manager placeholder VMs
            Write-PScriboMessage 'Creating VM lookup hashtable.'
            $VMs = Get-VM -Server $ESXi | Where-Object {
                $_.ExtensionData.Config.ManagedBy.ExtensionKey -notlike 'com.vmware.vcDr*'
            } | Sort-Object Name
            $VMLookup = @{ }
            foreach ($VM in $VMs) {
                $VMLookup.($VM.Id) = $VM.Name
            }

            # Create a lookup hashtable to link Host MoRefs to Names
            # Exclude VMware HCX hosts and ESX/ESXi versions prior to vSphere 5.0 from VMHost lookup
            Write-PScriboMessage 'Creating VMHost lookup hashtable.'
            $VMHost = Get-VMHost -Server $ESXi | Where-Object { $_.Model -notlike "*VMware Mobility Platform" -and $_.Version -gt 5 } | Sort-Object Name
            $VMHostLookup = @{ }
            $VMHostLookup.($VMHost.Id) = $VMHost.ExtensionData.Name

            # Create a lookup hashtable to link Datastore MoRefs to Names
            Write-PScriboMessage 'Creating Datastore lookup hashtable.'
            $Datastores = Get-Datastore -Server $ESXi | Where-Object { ($_.State -eq 'Available') -and ($_.CapacityGB -gt 0) } | Sort-Object Name
            $DatastoreLookup = @{ }
            foreach ($Datastore in $Datastores) {
                $DatastoreLookup.($Datastore.Id) = $Datastore.Name
            }

            # Create a lookup hashtable to link VDS Port Groups MoRefs to Names
            Write-PScriboMessage 'Creating VDPortGroup lookup hashtable.'
            $VDPortGroups = Get-VDPortgroup -Server $ESXi | Sort-Object Name
            $VDPortGroupLookup = @{ }
            foreach ($VDPortGroup in $VDPortGroups) {
                $VDPortGroupLookup.($VDPortGroup.Key) = $VDPortGroup.Name
            }

            Write-PScriboMessage "VMHost InfoLevel set at $($InfoLevel.VMHost)."
            #region Hosts Section
            if ($VMHost | Where-Object { $_.ConnectionState -eq 'Connected' -or $_.ConnectionState -eq 'Maintenance' }) {
                #region ESXi Host Detailed Information
                Section -Style Heading1 $($VMHost.ExtensionData.Name) {
                    Paragraph "The following sections detail the configuration of VMware ESXi host $($VMHost.ExtensionData.Name)."
                    # TODO: Host Certificate, Swap File Location
                    if ($InfoLevel.VMHost -ge 1) {
                        #region ESXi Host Hardware Section
                        Section -Style Heading2 'Hardware' {
                            Paragraph "The following section details the host hardware configuration for $($VMHost.ExtensionData.Name)."
                            BlankLine

                            #region ESXi Host Specifications
                            $VMHostUptime = Get-Uptime -VMHost $VMHost
                            $esxcli = Get-EsxCli -VMHost $VMHost -V2
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
                                'Serial Number' = Switch ($VMHost.ExtensionData.Hardware.SystemInfo.SerialNumber) {
                                    $null { '--' }
                                    default { $VMHost.ExtensionData.Hardware.SystemInfo.SerialNumber }
                                }
                                'Asset Tag' = Switch ($VMHost.ExtensionData.Summary.Hardware.OtherIdentifyingInfo[0].IdentifierValue) {
                                    '' { 'Unknown' }
                                    $null  { 'Unknown' }
                                    default { $VMHost.ExtensionData.Summary.Hardware.OtherIdentifyingInfo[0].IdentifierValue }
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
                                'Number of NICs' = $VMHost.ExtensionData.Summary.Hardware.NumNics
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
                            $TableParams = @{
                                Name = "ESXi Host Configuration - $($VMHost.ExtensionData.Name)"
                                List = $true
                                ColumnWidths = 50, 50
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $VMHostDetail | Table @TableParams
                            #endregion ESXi Host Specifications

                            #region ESXi IPMI/BMC Settings
                            Try {
                                $VMHostIPMI = $esxcli.hardware.ipmi.bmc.get.invoke()
                            } Catch {
                                Write-PScriboMessage -IsWarning "Unable to collect IPMI / BMC configuration from $($VMHost.ExtensionData.Name)"
                            }
                            if ($VMHostIPMI) {
                                Section -Style Heading3 'IPMI / BMC' {
                                    $VMHostIPMIInfo = [PSCustomObject]@{
                                        'Manufacturer' = $VMHostIPMI.Manufacturer
                                        'MAC Address' = $VMHostIPMI.MacAddress
                                        'IP Address' = $VMHostIPMI.IPv4Address
                                        'Subnet Mask' = $VMHostIPMI.IPv4Subnet
                                        'Gateway' = $VMHostIPMI.IPv4Gateway
                                        'Firmware Version' = $VMHostIPMI.BMCFirmwareVersion
                                    }

                                    $TableParams = @{
                                        Name = "IPMI / BMC - $($VMHost.ExtensionData.Name)"
                                        List = $true
                                        ColumnWidths = 50, 50
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VMHostIPMIInfo | Table @TableParams
                                }
                            }
                            #endregion ESXi IPMI/BMC Settings

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
                                $TableParams = @{
                                    Name = "Boot Device - $($VMHost.ExtensionData.Name)"
                                    List = $true
                                    ColumnWidths = 50, 50
                                }
                                if ($Report.ShowTableCaptions) {
                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                }
                                $VMHostBootDevice | Table @TableParams
                            }
                            #endregion ESXi Host Boot Devices

                            #region ESXi Host PCI Devices
                            Try {
                                Section -Style Heading3 'PCI Devices' {
                                    <# Move away from esxcli.v2 implementation to be compatible with 8.x branch.
                                    'Slot Description' information does not seem to be available through the API
                                    Create an array with PCI Address and VMware Devices (vmnic,vmhba,?vmgfx?)
                                    #>
                                    $PciToDeviceMapping = @{}
                                    $NetworkAdapters  = Get-VMHostNetworkAdapter -VMHost $VMHost -Physical
                                    foreach ($adapter in $NetworkAdapters) {
                                        $PciToDeviceMapping[$adapter.PciId] = $adapter.DeviceName
                                    }
                                    $hbAdapters = Get-VMHostHba -VMHost $VMHost
                                    foreach ($adapter in $hbAdapters) {
                                        $PciToDeviceMapping[$adapter.Pci] = $adapter.Device
                                    }
                                    <# Data Object - HostGraphicsInfo(vim.host.GraphicsInfo)
                                    This function has been available since version 5.5, but we can't be sure if it is still valid.
                                    I don't have access to a vGPU-enabled system.
                                    #>
                                    $GpuAdapters = (Get-VMHost $VMhost | Get-View -Property Config).Config.GraphicsInfo
                                    foreach ($adapter in $GpuAdapters) {
                                        $PciToDeviceMapping[$adapter.pciId] = $adapter.deviceName
                                    }

                                    $VMHostPciDevice = @{
                                        VMHost      = $VMHost
                                        DeviceClass = @('MassStorageController', 'NetworkController', 'DisplayController', 'SerialBusController')
                                    }
                                    $PciDevices = Get-VMHostPciDevice @VMHostPciDevice

                                    # Combine PciDevices and PciToDeviceMapping

                                    $VMHostPciDevices = $PciDevices | ForEach-Object {
                                        $PciDevice = $_
                                        $device = $PCIToDeviceMapping[$pciDevice.Id]

                                        if ($device) {
                                            [PSCustomObject]@{
                                                'Device'       = $device
                                                'PCI Address'   = $PciDevice.Id
                                                'Device Class'  = $PciDevice.DeviceClass -replace ('Controller',"")
                                                'Device Name'   = $PciDevice.DeviceName
                                                'Vendor Name'   = $PciDevice.VendorName
                                            }
                                        }
                                    }
                                    $TableParams = @{
                                        Name = "PCI Devices - $VMHost"
                                        ColumnWidths = 17, 18, 15, 30, 20
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VMHostPciDevices | Sort-Object 'Device' | Table @TableParams
                                }
                            } Catch {Write-PScriboMessage -IsWarning "Unable to collect PCI Devices information from $VMHost"}
                            #endregion ESXi Host PCI Devices

                            #region ESXi Host PCI Devices Drivers & Firmware
                            Try {
                                Section -Style Heading3 'PCI Devices Drivers & Firmware' {
                                    $PciToDeviceMapping = @{}
                                    $NetworkAdapters  = Get-VMHostNetworkAdapter -VMHost $VMHost -Physical
                                    foreach ($adapter in $NetworkAdapters) {
                                        $PciToDeviceMapping[$adapter.PciId] = $adapter.DeviceName
                                    }
                                    $hbAdapters = Get-VMHostHba -VMHost $VMHost
                                    foreach ($adapter in $hbAdapters) {
                                        $PciToDeviceMapping[$adapter.Pci] = $adapter.Device
                                    }
                                    <# Data Object - HostGraphicsInfo(vim.host.GraphicsInfo)
                                    This function has been available since version 5.5, but we can't be sure if it is still valid.
                                    I don't have access to a vGPU-enabled system.
                                    #>
                                    $GpuAdapters = (Get-VMHost $VMhost | Get-View -Property Config).Config.GraphicsInfo
                                    foreach ($adapter in $GpuAdapters) {
                                        $PciToDeviceMapping[$adapter.pciId] = $adapter.deviceName
                                    }

                                    $VMHostPciDevice = @{
                                        VMHost      = $VMHost
                                        DeviceClass = @('MassStorageController', 'NetworkController', 'DisplayController', 'SerialBusController')
                                    }
                                    $PciDevices = Get-VMHostPciDevice @VMHostPciDevice

                                    # Combine PciDevices and PciToDeviceMapping

                                    $VMHostPciDevicesDetails = $PciDevices | ForEach-Object {
                                        $PciDevice = $_
                                        $device = $PCIToDeviceMapping[$pciDevice.Id]

                                        if ($device) {
                                            [PSCustomObject]@{
                                                'Device' = $device
                                                'Model' = $PciDevice.DeviceName
                                                'Driver' = Switch ($PciDevice.DeviceClass) {
                                                    'NetworkController' {($NetworkAdapters.ExtensionData | Where-Object {$_.Pci -eq $PciDevice.Id}).Driver}
                                                    'MassStorageController' {($hbAdapters.ExtensionData | Where-Object {$_.Pci -eq $PciDevice.Id}).Driver}
                                                    default {'--'}
                                                }
                                                'Driver Version' = Switch ($PciDevice.DeviceClass) {
                                                    'NetworkController' {$esxcli.system.module.get.Invoke(@{module = ($NetworkAdapters.ExtensionData | Where-Object {$_.Pci -eq $PciDevice.Id}).Driver }).Version}
                                                    'MassStorageController' {$esxcli.system.module.get.Invoke(@{module = ($hbAdapters.ExtensionData | Where-Object {$_.Pci -eq $PciDevice.Id}).Driver}).Version}
                                                    default {'--'}
                                                }
                                                'Firmware Version' = Switch ($PciDevice.DeviceClass) {
                                                    'NetworkController' {$esxcli.network.nic.get.Invoke(@{ nicname = $device }).DriverInfo.FirmwareVersion}
                                                    default {'--'}
                                                }
                                                'VIB Name' = Switch ($PciDevice.DeviceClass) {
                                                    'NetworkController' {($esxcli.software.vib.list.Invoke() | Select-Object -Property Name, Version | Where-Object { $_.Name -eq (($NetworkAdapters.ExtensionData | Where-Object {$_.Pci -eq $PciDevice.Id}).Driver) -or $_.Name -eq "net-" + (($NetworkAdapters.ExtensionData | Where-Object {$_.Pci -eq $PciDevice.Id}).Driver) -or $_.Name -eq "net55-" + (($NetworkAdapters.ExtensionData | Where-Object {$_.Pci -eq $PciDevice.Id}).Driver) }).Name}
                                                    'MassStorageController' {($esxcli.software.vib.list.Invoke() | Select-Object -Property Name, Version | Where-Object { $_.Name -eq "scsi-" + (($hbAdapters.ExtensionData | Where-Object {$_.Pci -eq $PciDevice.Id}).Driver -replace "_", "-") -or $_.Name -eq "sata-" + (($hbAdapters.ExtensionData | Where-Object {$_.Pci -eq $PciDevice.Id}).Driver -replace "_", "-") -or $_.Name -eq (($hbAdapters.ExtensionData | Where-Object {$_.Pci -eq $PciDevice.Id}).Driver -replace "_", "-")}).Name}
                                                    default {'--'}
                                                }
                                                'VIB Version' = Switch ($PciDevice.DeviceClass) {
                                                    'NetworkController' {($esxcli.software.vib.list.Invoke() | Select-Object -Property Name, Version | Where-Object { $_.Name -eq (($NetworkAdapters.ExtensionData | Where-Object {$_.Pci -eq $PciDevice.Id}).Driver) -or $_.Name -eq "net-" + (($NetworkAdapters.ExtensionData | Where-Object {$_.Pci -eq $PciDevice.Id}).Driver) -or $_.Name -eq "net55-" + (($NetworkAdapters.ExtensionData | Where-Object {$_.Pci -eq $PciDevice.Id}).Driver) }).Version}
                                                    'MassStorageController' {($esxcli.software.vib.list.Invoke() | Select-Object -Property Name, Version | Where-Object { $_.Name -eq "scsi-" + (($hbAdapters.ExtensionData | Where-Object {$_.Pci -eq $PciDevice.Id}).Driver -replace "_", "-") -or $_.Name -eq "sata-" + (($hbAdapters.ExtensionData | Where-Object {$_.Pci -eq $PciDevice.Id}).Driver -replace "_", "-") -or $_.Name -eq (($hbAdapters.ExtensionData | Where-Object {$_.Pci -eq $PciDevice.Id}).Driver -replace "_", "-")}).Version}
                                                    default {'--'}
                                                }
                                            }
                                        }
                                    }
                                    $TableParams = @{
                                        Name = "PCI Devices Drivers & Firmware - $VMHost"
                                        ColumnWidths = 12, 20, 11, 19, 11, 11, 16
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VMHostPciDevicesDetails | Sort-Object 'Device' | Table @TableParams
                                }
                            } Catch {Write-PScriboMessage -IsWarning "Unable to collect PCI Devices Drivers & Firmware information from $VMHost"}
                            #endregion ESXi Host PCI Devices Drivers & Firmware
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
                                $TableParams = @{
                                    Name = "Image Profile - $($VMHost.ExtensionData.Name)"
                                    #ColumnWidths = 50, 25, 25
                                }
                                if ($Report.ShowTableCaptions) {
                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                }
                                $SecurityProfile | Table @TableParams
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
                                $TableParams = @{
                                    Name = "Time Configuration - $($VMHost.ExtensionData.Name)"
                                    ColumnWidths = 30, 30, 40
                                }
                                if ($Report.ShowTableCaptions) {
                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                }
                                $VMHostTimeSettings | Table @TableParams
                            }
                            #endregion ESXi Host Time Configuration

                            #region ESXi Host Syslog Configuration
                            $SyslogConfig = $VMHost | Get-VMHostSysLogServer
                            if ($SyslogConfig) {
                                Section -Style Heading3 'Syslog Configuration' {
                                    # TODO: Syslog Rotate & Size, Log Directory (Adv Settings)
                                    $SyslogConfig = $SyslogConfig | Select-Object @{L = 'SysLog Server'; E = { $_.Host } }, Port
                                    $TableParams = @{
                                        Name = "Syslog Configuration - $($VMHost.ExtensionData.Name)"
                                        ColumnWidths = 50, 50
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $SyslogConfig | Table @TableParams
                                }
                            }
                            #endregion ESXi Host Syslog Configuration

                            #region ESXi Host Comprehensive Information Section
                            if ($InfoLevel.VMHost -ge 5) {
                                #region ESXi Host Advanced System Settings
                                Section -Style Heading3 'Advanced System Settings' {
                                    $AdvSettings = $VMHost | Get-AdvancedSetting | Select-Object Name, Value
                                    $TableParams = @{
                                        Name = "Advanced System Settings - $($VMHost.ExtensionData.Name)"
                                        ColumnWidths = 50, 50
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $AdvSettings | Sort-Object Name | Table @TableParams
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
                                    $TableParams = @{
                                        Name = "Software VIBs - $($VMHost.ExtensionData.Name)"
                                        ColumnWidths = 15, 25, 15, 15, 15, 15
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VMHostVibs | Sort-Object 'Install Date' -Descending | Table @TableParams
                                }
                                #endregion ESXi Host Software VIBs
                            }
                            #endregion ESXi Host Comprehensive Information Section
                        }
                        #endregion ESXi Host System Section

                        #region ESXi Host Storage Section
                        if ($InfoLevel.Storage -ge 1) {
                            Section -Style Heading2 'Storage' {
                                Paragraph "The following section details the host storage configuration for $($VMHost.ExtensionData.Name)."

                                #region Datastore Section
                                Write-PScriboMessage "Storage InfoLevel set at $($InfoLevel.Storage)."

                                if ($Datastores) {
                                    Section -Style Heading3 'Datastores' {
                                        #region Datastore Infomative Information
                                        if (($InfoLevel.Storage -ge 1) -and ($InfoLevel.Storage -lt 3)) {
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
                                            $TableParams = @{
                                                Name = "Datastores - $($VMHost.ExtensionData.Name)"
                                                ColumnWidths = 20, 8, 9, 8, 15, 15, 15, 10
                                            }
                                            if ($Report.ShowTableCaptions) {
                                                $TableParams['Caption'] = "- $($TableParams.Name)"
                                            }
                                            $DatastoreInfo | Sort-Object 'Datastore' | Table @TableParams
                                        }
                                        #endregion Datastore Advanced Summary

                                        #region Datastore Detailed Information
                                        if ($InfoLevel.Storage -ge 3) {
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
                                                    if ($InfoLevel.Storage -ge 4) {
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
                                                    $TableParams = @{
                                                        Name = "Datastore $($Datastore.Name) - $($VMHost.ExtensionData.Name)"
                                                        List = $true
                                                        ColumnWidths = 50, 50
                                                    }
                                                    if ($Report.ShowTableCaptions) {
                                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                                    }
                                                    $DatastoreDetail | Sort-Object Datacenter, Datastore | Table @TableParams

                                                    # Get VMFS volumes. Ignore local SCSILuns.
                                                    if (($Datastore.Type -eq 'VMFS') -and ($Datastore.ExtensionData.Info.Vmfs.Local -eq $false)) {
                                                        #region SCSI LUN Information Section
                                                        Section -Style Heading4 'SCSI LUNs' {
                                                            $ScsiLuns = foreach ($DatastoreHost in $Datastore.ExtensionData.Host.Key) {
                                                                $DiskName = $Datastore.ExtensionData.Info.Vmfs.Extent.DiskName
                                                                $ScsiDeviceDetailProps = @{
                                                                    'VMHosts' = $VMHost
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
                                                            $TableParams = @{
                                                                Name = "SCSI LUNs - $($VMHost.ExtensionData.Name)"
                                                                ColumnWidths = 18, 18, 10, 14, 12, 8, 12, 8
                                                            }
                                                            if ($Report.ShowTableCaptions) {
                                                                $TableParams['Caption'] = "- $($TableParams.Name)"
                                                            }
                                                            $ScsiLuns | Sort-Object Host | Table @TableParams
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
                                #endregion Datastore Section

                                #region ESXi Host Storage Adapter Information
                                $VMHostHbas = $VMHost | Get-VMHostHba | Sort-Object Device
                                if ($VMHostHbas) {
                                    #region ESXi Host Storage Adapters Section
                                    Section -Style Heading3 'Storage Adapters' {
                                        if ($InfoLevel.VMHost -ge 3) {
                                            foreach ($VMHostHba in $VMHostHbas) {
                                                $Target = ((Get-View $VMHostHba.VMhost).Config.StorageDevice.ScsiTopology.Adapter | Where-Object { $_.Adapter -eq $VMHostHba.Key }).Target
                                                $LUNs = Get-ScsiLun -Hba $VMHostHba -LunType "disk" -ErrorAction SilentlyContinue
                                                $Paths = ($Target | ForEach-Object { $_.Lun.Count } | Measure-Object -Sum)
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
                                                    $TableParams = @{
                                                        Name = "Storage Adapter $($VMHostStorageAdapter.Adapter) - $($VMHost.ExtensionData.Name)"
                                                        List = $true
                                                        ColumnWidths = 25, 75
                                                    }
                                                    if ($Report.ShowTableCaptions) {
                                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                                    }
                                                    $VMHostStorageAdapter | Table @TableParams
                                                }
                                            }
                                        } else {
                                            $VMHostStorageAdapters = foreach ($VMHostHba in $VMHostHbas) {
                                                [PSCustomObject]@{
                                                    'Adapter' = $VMHostHba.Device
                                                    'Type' = Switch ($VMHostHba.Type) {
                                                        'FibreChannel' { 'Fibre Channel' }
                                                        'IScsi' { 'iSCSI' }
                                                        'ParallelScsi' { 'Parallel SCSI' }
                                                        default { $TextInfo.ToTitleCase($VMHostHba.Type) }
                                                    }
                                                    'Model' = $VMHostHba.Model
                                                    'Status' = $TextInfo.ToTitleCase($VMHostHba.Status)
                                                }
                                            }
                                            if ($Healthcheck.VMHost.StorageAdapter) {
                                                $VMHostStorageAdapters | Where-Object { $_.'Status' -ne 'Online' } | Set-Style -Style Warning -Property 'Status'
                                                $VMHostStorageAdapters | Where-Object { $_.'Status' -eq 'Offline' } | Set-Style -Style Critical -Property 'Status'
                                            }
                                            $TableParams = @{
                                                Name = "Storage Adapters - $($VMHost.ExtensionData.Name)"
                                                ColumnWidths = 25, 25, 25, 25
                                            }
                                            if ($Report.ShowTableCaptions) {
                                                $TableParams['Caption'] = "- $($TableParams.Name)"
                                            }
                                            $VMHostStorageAdapters | Table @TableParams
                                        }
                                    }
                                    #endregion ESXi Host Storage Adapters Section
                                }
                                #endregion ESXi Host Storage Adapter Information
                            }
                        }
                        #endregion ESXi Host Storage Section
                    }

                    #region ESXi Host Network Section
                    if ($InfoLevel.Network -ge 1) {
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
                            $TableParams = @{
                                Name = "Network Configuration - $($VMHost.ExtensionData.Name)"
                                List = $true
                                ColumnWidths = 50, 50
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $VMHostNetworkDetail | Table @TableParams
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
                                            $TableParams = @{
                                                Name = "Physical Adapter $($VMHostPhysicalNetAdapter.Adapter) - $($VMHost.ExtensionData.Name)"
                                                List = $true
                                                ColumnWidths = 50, 50
                                            }
                                            if ($Report.ShowTableCaptions) {
                                                $TableParams['Caption'] = "- $($TableParams.Name)"
                                            }
                                            $VMHostPhysicalNetAdapter | Table @TableParams
                                        }
                                    }
                                } else {
                                    $TableParams = @{
                                        Name = "Physical Adapters - $($VMHost.ExtensionData.Name)"
                                        ColumnWidths = 11, 13, 15, 19, 14, 14, 14
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VMHostPhysicalNetAdapters | Table @TableParams
                                }
                            }
                            #endregion ESXi Host Physical Adapters

                            #region ESXi Host Cisco Discovery Protocol
                            $VMHostNetworkAdapterCDP = $VMHost | Get-VMHostNetworkAdapterDP | Where-Object { $_.Status -eq 'Connected' } | Sort-Object Device
                            if ($VMHostNetworkAdapterCDP) {
                                Section -Style Heading3 'Cisco Discovery Protocol' {
                                    if ($InfoLevel.VMHost -ge 4) {
                                        foreach ($VMHostNetworkAdapter in $VMHostNetworkAdapterCDP) {
                                            Section -Style Heading5 "$($VMHostNetworkAdapter.Device)" {
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
                                                $TableParams = @{
                                                    Name = "Network Adapter $($VMHostNetworkAdapter.Device) CDP Information - $($VMHost.ExtensionData.Name)"
                                                    List = $true
                                                    ColumnWidths = 50, 50
                                                }
                                                if ($Report.ShowTableCaptions) {
                                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                                }
                                                $VMHostCDP | Table @TableParams
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
                                        $TableParams = @{
                                            Name = "Network Adapter CDP Information - $($VMHost.ExtensionData.Name)"
                                            ColumnWidths = 11, 13, 26, 22, 17, 11
                                        }
                                        if ($Report.ShowTableCaptions) {
                                            $TableParams['Caption'] = "- $($TableParams.Name)"
                                        }
                                        $VMHostCDP | Table @TableParams
                                    }
                                }
                            }
                            #endregion ESXi Host Cisco Discovery Protocol

                            #region ESXi Host Link Layer Discovery Protocol
                            $VMHostNetworkAdapterLLDP = $VMHost | Get-VMHostNetworkAdapterDP | Where-Object { $null -ne $_.ChassisId } | Sort-Object Device
                            if ($VMHostNetworkAdapterLLDP) {
                                Section -Style Heading3 'Link Layer Discovery Protocol' {
                                    if ($InfoLevel.VMHost -ge 4) {
                                        foreach ($VMHostNetworkAdapter in $VMHostNetworkAdapterLLDP) {
                                            Section -Style Heading5 "$($VMHostNetworkAdapter.Device)" {
                                                $VMHostLLDP = [PSCustomObject]@{
                                                    'Chassis ID' = $VMHostNetworkAdapter.ChassisId
                                                    'Port ID' = $VMHostNetworkAdapter.PortId
                                                    'Time to live' = $VMHostNetworkAdapter.TimeToLive
                                                    'TimeOut' = $VMHostNetworkAdapter.TimeOut
                                                    'Samples' = $VMHostNetworkAdapter.Samples
                                                    'Management Address' = $VMHostNetworkAdapter.ManagementAddress
                                                    'Port Description' = $VMHostNetworkAdapter.PortDescription
                                                    'System Description' = $VMHostNetworkAdapter.SystemDescription
                                                    'System Name' = $VMHostNetworkAdapter.SystemName
                                                }
                                                $TableParams = @{
                                                    Name = "Network Adapter $($VMHostNetworkAdapter.Device) LLDP Information - $($VMHost.ExtensionData.Name)"
                                                    List = $true
                                                    ColumnWidths = 50, 50
                                                }
                                                if ($Report.ShowTableCaptions) {
                                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                                }
                                                $VMHostLLDP | Table @TableParams
                                            }
                                        }
                                    } else {
                                        $VMHostLLDP = foreach ($VMHostNetworkAdapter in $VMHostNetworkAdapterLLDP) {
                                            [PSCustomObject]@{
                                                'Adapter' = $VMHostNetworkAdapter.Device
                                                'Chassis ID' = $VMHostNetworkAdapter.ChassisId
                                                'Port ID' = $VMHostNetworkAdapter.PortId
                                                'Management Address' = $VMHostNetworkAdapter.ManagementAddress
                                                'Port Description' = $VMHostNetworkAdapter.PortDescription
                                                'System Name' = $VMHostNetworkAdapter.SystemName
                                            }
                                        }
                                        $TableParams = @{
                                            Name = "Network Adapter LLDP Information - $($VMHost.ExtensionData.Name)"
                                            ColumnWidths = 11, 19, 16, 19, 18, 17
                                        }
                                        if ($Report.ShowTableCaptions) {
                                            $TableParams['Caption'] = "- $($TableParams.Name)"
                                        }
                                        $VMHostLLDP | Table @TableParams
                                    }
                                }
                            }
                            #endregion ESXi Host Link Layer Discovery Protocol

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
                                                    (Get-VDPortgroup -Name $script:pg).VDSwitch.Name | Select-Object -Unique
                                                }
                                            }
                                            'TCP/IP Stack' = Switch ($_.Spec.NetstackInstanceKey) {
                                                'defaultTcpipStack' { 'Default' }
                                                'vSphereProvisioning' { 'Provisioning' }
                                                'vmotion' { 'vMotion' }
                                                'vxlan' { 'nsx-overlay' }
                                                'hyperbus' { 'nsx-hyperbus' }
                                                $null { 'Not Applicable' }
                                                default { $_.Spec.NetstackInstanceKey }
                                            }
                                            'Enabled Services' = Switch ( $vnicMgr.Info.NetConfig | Where-Object { $_.SelectedVnic -match $device } | ForEach-Object { $_.NicType } ) {
                                                'vmotion' { 'vMotion' }
                                                'vSphereProvisioning' { 'Provisioning' }
                                                'faultToleranceLogging' { 'FT Logging' }
                                                'management' { 'Management' }
                                                'vSphereReplication' { 'vSphere Replication' }
                                                'vSphereReplicationNFC' { 'vSphere Replication NFC' }
                                                'vsan' { 'vSAN' }
                                                'vsanWitness' { 'vSAN Witness' }
                                            }
                                            'MTU' = $_.Spec.Mtu
                                            'MAC Address' = $_.Spec.Mac
                                            'DHCP' = Switch ($_.Spec.Ip.Dhcp) {
                                                $true { 'Enabled' }
                                                $false { 'Disabled' }
                                            }
                                            'IP Address' = & {
                                                if ($_.Spec.IP.IPAddress) {
                                                    $script:ip = $_.Spec.IP.IPAddress
                                                } else {
                                                    $script:ip = '--'
                                                }
                                                $script:ip
                                            }
                                            'Subnet Mask' = & {
                                                if ($_.Spec.IP.SubnetMask) {
                                                    $script:netmask = $_.Spec.IP.SubnetMask
                                                } else {
                                                    $script:netmask = '--'
                                                }
                                                $script:netmask
                                            }
                                            'Default Gateway' = Switch ($_.Spec.IpRouteSpec.IpRouteConfig.DefaultGateway) {
                                                $null { '--' }
                                                default { $_.Spec.IpRouteSpec.IpRouteConfig.DefaultGateway }
                                            }
                                        }
                                    }
                                }

                                if ($InfoLevel.VMHost -ge 3) {
                                    foreach ($VMkernelAdapter in ($VMkernelAdapters | Sort-Object 'Adapter')) {
                                        Section -Style Heading4 "$($VMkernelAdapter.Adapter)" {
                                            $TableParams = @{
                                                Name = "VMkernel Adapter $($VMkernelAdapter.Adapter) - $($VMHost.ExtensionData.Name)"
                                                List = $true
                                                ColumnWidths = 50, 50
                                            }
                                            if ($Report.ShowTableCaptions) {
                                                $TableParams['Caption'] = "- $($TableParams.Name)"
                                            }
                                            $VMkernelAdapter | Table @TableParams
                                        }
                                    }
                                } else {
                                    $TableParams = @{
                                        Name = "VMkernel Adapters - $($VMHost.ExtensionData.Name)"
                                        Columns = 'Adapter', 'Port Group', 'TCP/IP Stack', 'Enabled Services','IP Address'
                                        ColumnWidths = 11, 35, 18, 18, 18
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VMkernelAdapters | Sort-Object 'Adapter' | Table @TableParams
                                }
                            }
                            #endregion ESXi Host VMkernel Adapaters

                            #region ESXi Host Standard Virtual Switches
                            $VSSwitches = $VMHost | Get-VirtualSwitch -Standard | Sort-Object Name
                            if ($VSSwitches) {
                                #region Section Standard Virtual Switches
                                Section -Style Heading5 'Standard Virtual Switches' {
                                    Paragraph "The following section details the standard virtual switch configuration for $($VMHost.ExtensionData.Name)."
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
                                    $TableParams = @{
                                        Name = "Standard Virtual Switches - $($VMHost.ExtensionData.Name)"
                                        ColumnWidths = 25, 25, 25, 25
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VSSProperties | Table @TableParams
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
                                            $TableParams = @{
                                                Name = "Virtual Switch Security Policy - $($VMHost.ExtensionData.Name)"
                                                ColumnWidths = 25, 25, 25, 25
                                            }
                                            if ($Report.ShowTableCaptions) {
                                                $TableParams['Caption'] = "- $($TableParams.Name)"
                                            }
                                            $VssSecurity | Sort-Object 'Virtual Switch' | Table @TableParams
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
                                        $TableParams = @{
                                            Name = "Virtual Switch Traffic Shaping Policy - $($VMHost.ExtensionData.Name)"
                                            ColumnWidths = 25, 15, 20, 20, 20
                                        }
                                        if ($Report.ShowTableCaptions) {
                                            $TableParams['Caption'] = "- $($TableParams.Name)"
                                        }
                                        $VssTrafficShapingPolicy | Sort-Object 'Virtual Switch' | Table @TableParams
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
                                            $TableParams = @{
                                                Name = "Virtual Switch Teaming & Failover - $($VMHost.ExtensionData.Name)"
                                                ColumnWidths = 20, 17, 12, 11, 10, 10, 10, 10
                                            }
                                            if ($Report.ShowTableCaptions) {
                                                $TableParams['Caption'] = "- $($TableParams.Name)"
                                            }
                                            $VssNicTeaming | Sort-Object 'Virtual Switch' | Table @TableParams
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
                                            $TableParams = @{
                                                Name = "Virtual Switch Port Group Information - $($VMHost.ExtensionData.Name)"
                                                ColumnWidths = 40, 10, 40, 10
                                            }
                                            if ($Report.ShowTableCaptions) {
                                                $TableParams['Caption'] = "- $($TableParams.Name)"
                                            }
                                            $VssPortgroups | Sort-Object 'Port Group', 'VLAN ID', 'Virtual Switch' | Table @TableParams
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
                                                $TableParams = @{
                                                    Name = "Virtual Switch Port Group Security Policy - $($VMHost.ExtensionData.Name)"
                                                    ColumnWidths = 27, 25, 16, 16, 16
                                                }
                                                if ($Report.ShowTableCaptions) {
                                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                                }
                                                $VssPortgroupSecurity | Sort-Object 'Port Group', 'Virtual Switch' | Table @TableParams
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
                                            $TableParams = @{
                                                Name = "Virtual Switch Port Group Traffic Shaping Policy - $($VMHost.ExtensionData.Name)"
                                                ColumnWidths = 19, 19, 11, 17, 17, 17
                                            }
                                            if ($Report.ShowTableCaptions) {
                                                $TableParams['Caption'] = "- $($TableParams.Name)"
                                            }
                                            $VssPortgroupTrafficShapingPolicy | Sort-Object 'Port Group', 'Virtual Switch' | Table @TableParams
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
                                                $TableParams = @{
                                                    Name = "Virtual Switch Port Group Teaming & Failover - $($VMHost.ExtensionData.Name)"
                                                    ColumnWidths = 12, 11, 11, 11, 11, 11, 11, 11, 11
                                                }
                                                if ($Report.ShowTableCaptions) {
                                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                                }
                                                $VssPortgroupNicTeaming | Sort-Object 'Port Group', 'Virtual Switch' | Table @TableParams
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
                                    #region Distributed Virtual Switch Advanced Summary
                                    if ($InfoLevel.Network -le 2) {
                                        $VDSInfo = foreach ($VDS in $VDSwitches) {
                                            [PSCustomObject]@{
                                                'Distributed Switch' = $VDS.Name
                                                'Number of Ports' = $VDS.NumPorts
                                                'Number of Port Groups' = ($VDS.ExtensionData.Summary.PortGroupName).Count
                                                'MTU' = $VDS.Mtu
                                                'Discovery Protocol Type' = $VDS.LinkDiscoveryProtocol
                                                'Discovery Protocol Operation' = $VDS.LinkDiscoveryProtocolOperation
                                            }
                                        }
                                        $TableParams = @{
                                            Name = "Distributed Switch Information - $($VMHost.ExtensionData.Name)"
                                            ColumnWidths = 25, 15, 15, 15, 15, 15
                                        }
                                        if ($Report.ShowTableCaptions) {
                                            $TableParams['Caption'] = "- $($TableParams.Name)"
                                        }
                                        $VDSInfo | Table @TableParams
                                    }
                                    #endregion Distributed Switch Advanced Summary

                                    #region Distributed Switch Detailed Information
                                    if ($InfoLevel.Network -ge 3) {
                                        # TODO: LACP, NetFlow, NIOC
                                        foreach ($VDS in ($VDSwitches)) {
                                            $VdsVmCount = ($VDS | Get-VM).Count
                                            #region VDS Section
                                            Section -Style Heading4 $VDS {
                                                #region Distributed Switch General Properties
                                                $VDSwitchDetail = [PSCustomObject]@{
                                                    'Distributed Switch' = $VDS.Name
                                                    'ID' = $VDS.Id
                                                    'Number of Ports' = $VDS.NumPorts
                                                    'Number of Port Groups' = ($VDS.ExtensionData.Summary.PortGroupName).Count
                                                    'Number of VMs' = $VdsVmCount
                                                    'MTU' = $VDS.Mtu
                                                    'Network I/O Control' = Switch ($VDS.ExtensionData.Config.NetworkResourceManagementEnabled) {
                                                        $true { 'Enabled' }
                                                        $false { 'Disabled' }
                                                    }
                                                    'Discovery Protocol' = $VDS.LinkDiscoveryProtocol
                                                    'Discovery Protocol Operation' = $VDS.LinkDiscoveryProtocolOperation
                                                }
                                                # TODO: Fix this, incorrect reporting!
                                                #region Network Advanced Detail Information
                                                if ($InfoLevel.Network -ge 4) {
                                                    $VDSwitchVMs = $VDS | Get-VM | Sort-Object
                                                    Add-Member -InputObject $VDSwitchDetail -MemberType NoteProperty -Name 'Virtual Machines' -Value ($VDSwitchVMs.Name -join ', ')
                                                }
                                                #endregion Network Advanced Detail Information
                                                $TableParams = @{
                                                    Name = "$VDS Distributed Switch General Properties - $($VMHost.ExtensionData.Name)"
                                                    List = $true
                                                    ColumnWidths = 50, 50
                                                }
                                                if ($Report.ShowTableCaptions) {
                                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                                }
                                                $VDSwitchDetail | Table @TableParams
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
                                                        $TableParams = @{
                                                            Name = "$VDS Distributed Switch Uplink Ports - $($VMHost.ExtensionData.Name)"
                                                            ColumnWidths = 30, 20, 20, 30
                                                        }
                                                        if ($Report.ShowTableCaptions) {
                                                            $TableParams['Caption'] = "- $($TableParams.Name)"
                                                        }
                                                        $VdsUplinkDetail | Sort-Object 'Distributed Switch', 'Uplink Name' | Table @TableParams
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
                                                        $TableParams = @{
                                                            Name = "$VDS Distributed Switch Port Groups - $($VMHost.ExtensionData.Name)"
                                                            ColumnWidths = 35, 35, 15, 15
                                                        }
                                                        if ($Report.ShowTableCaptions) {
                                                            $TableParams['Caption'] = "- $($TableParams.Name)"
                                                        }
                                                        $VDSPortgroupDetail | Sort-Object 'Port Group' | Table @TableParams
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
                                                        $TableParams = @{
                                                            Name = "$VDS Distributed Switch Private VLANs - $($VMHost.ExtensionData.Name)"
                                                            ColumnWidths = 33, 34, 33
                                                        }
                                                        if ($Report.ShowTableCaptions) {
                                                            $TableParams['Caption'] = "- $($TableParams.Name)"
                                                        }
                                                        $VDSPvlan | Sort-Object 'Primary VLAN ID', 'Secondary VLAN ID' | Table @TableParams
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
                    }
                    #endregion ESXi Host Network Section

                    #region ESXi Host Security Section
                    if ($InfoLevel.VMHost -ge 1) {
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
                                    $TableParams = @{
                                        Name = "Lockdown Mode - $($VMHost.ExtensionData.Name)"
                                        List = $true
                                        ColumnWidths = 50, 50
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $LockdownMode | Table @TableParams
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
                                $TableParams = @{
                                    Name = "Services - $($VMHost.ExtensionData.Name)"
                                    ColumnWidths = 40, 20, 40
                                }
                                if ($Report.ShowTableCaptions) {
                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                }
                                $Services | Sort-Object 'Service' | Table @TableParams
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
                                        $TableParams = @{
                                            Name = "Firewall Configuration - $($VMHost.ExtensionData.Name)"
                                            ColumnWidths = 22, 12, 21, 21, 12, 12
                                        }
                                        if ($Report.ShowTableCaptions) {
                                            $TableParams['Caption'] = "- $($TableParams.Name)"
                                        }
                                        $VMHostFirewall | Sort-Object 'Service' | Table @TableParams
                                    }
                                    #endregion Friewall Section
                                }
                                #endregion ESXi Host Firewall

                                #region ESXi Host Authentication
                                $AuthServices = $VMHost | Get-VMHostAuthentication
                                if ($AuthServices.DomainMembershipStatus) {
                                    Section -Style Heading3 'Authentication Services' {
                                        $AuthServices = $AuthServices | Select-Object Domain, @{L = 'Domain Membership'; E = { $_.DomainMembershipStatus } }, @{L = 'Trusted Domains'; E = { $_.TrustedDomains } }
                                        $TableParams = @{
                                            Name = "Authentication Services - $($VMHost.ExtensionData.Name)"
                                            ColumnWidths  = 25, 25, 50
                                        }
                                        if ($Report.ShowTableCaptions) {
                                            $TableParams['Caption'] = "- $($TableParams.Name)"
                                        }
                                        $AuthServices | Table @TableParams
                                    }
                                }
                                #endregion ESXi Host Authentication
                            }
                            #endregion ESXi Host Advanced Detail Information
                        }
                    }
                    #endregion ESXi Host Security Section

                    #region Virtual Machine Section
                    Write-PScriboMessage "VM InfoLevel set at $($InfoLevel.VM)."
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
                                    $TableParams = @{
                                        Name = "VM Summary - $($VMHost.ExtensionData.Name)"
                                        List = $true
                                        ColumnWidths  = 50, 50
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VMSummary | Table @TableParams
                                }
                                #endregion Virtual Machine Summary Information

                                #region Virtual Machine Advanced Summary
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
                                    $TableParams = @{
                                        Name = "VM Advanced Summary - $($VMHost.ExtensionData.Name)"
                                        ColumnWidths = 21, 8, 16, 9, 9, 9, 9, 9, 10
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VMInfo | Table @TableParams

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
                                            $TableParams = @{
                                                Name = "VM Snapshot Information - $($VMHost.ExtensionData.Name)"
                                                ColumnWidths = 30, 30, 30, 10
                                            }
                                            if ($Report.ShowTableCaptions) {
                                                $TableParams['Caption'] = "- $($TableParams.Name)"
                                            }
                                            $VMSnapshotInfo | Table @TableParams
                                        }
                                    }
                                    #endregion VM Snapshot Information
                                }
                                #endregion Virtual Machine Advanced Summary

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
                                            $TableParams = @{
                                                Name = "$($VM.Name) VM Configuration - $($VMHost.ExtensionData.Name)"
                                                List = $true
                                                ColumnWidths = 50, 50
                                            }
                                            if ($Report.ShowTableCaptions) {
                                                $TableParams['Caption'] = "- $($TableParams.Name)"
                                            }
                                            $VMDetail | Table @TableParams

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
                                                        $TableParams = @{
                                                            Name = "$($VM.Name) Network Adapters - $($VMHost.ExtensionData.Name)"
                                                            ColumnWidths = 20, 12, 16, 12, 20, 20
                                                        }
                                                        if ($Report.ShowTableCaptions) {
                                                            $TableParams['Caption'] = "- $($TableParams.Name)"
                                                        }
                                                        $VMnicInfo | Table @TableParams
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
                                                        $TableParams = @{
                                                            Name = "$($VM.Name) SCSI Controllers - $($VMHost.ExtensionData.Name)"
                                                            ColumnWidths = 33, 34, 33
                                                        }
                                                        if ($Report.ShowTableCaptions) {
                                                            $TableParams['Caption'] = "- $($TableParams.Name)"
                                                        }
                                                        $VMScsiControllers | Sort-Object 'Device' | Table @TableParams
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
                                                            $TableParams = @{
                                                                Name = "$($VM.Name) Hard Disks - $($VMHost.ExtensionData.Name)"
                                                                ColumnWidths = 15, 25, 15, 15, 15, 15
                                                            }
                                                            if ($Report.ShowTableCaptions) {
                                                                $TableParams['Caption'] = "- $($TableParams.Name)"
                                                            }
                                                            $VMHardDiskInfo | Table @TableParams
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
                                                                    $TableParams = @{
                                                                        Name = "$($VM.Name) $($VMHdd.Name) HDD Configuration - $($VMHost.ExtensionData.Name)"
                                                                        List = $true
                                                                        ColumnWidths = 25, 75
                                                                    }
                                                                    if ($Report.ShowTableCaptions) {
                                                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                                                    }
                                                                    $VMHardDiskInfo | Table @TableParams
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
                                                        $TableParams = @{
                                                            Name = "$($VM.Name) Guest Volumes - $($VMHost.ExtensionData.Name)"
                                                            ColumnWidths = 25, 25, 25, 25
                                                        }
                                                        if ($Report.ShowTableCaptions) {
                                                            $TableParams['Caption'] = "- $($TableParams.Name)"
                                                        }
                                                        $VMGuestDiskInfo | Table @TableParams
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
                                                    $TableParams = @{
                                                        Name = "$($VM.Name) VM Snapshots - $($VMHost.ExtensionData.Name)"
                                                        ColumnWidths = 45, 45, 10
                                                    }
                                                    if ($Report.ShowTableCaptions) {
                                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                                    }
                                                    $VMSnapshots | Table @TableParams
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

                    #region ESXi Host VM Autostart Information
                    if ($InfoLevel.VMHost -ge 1) {
                        $VMStartPolicy = $VMHost | Get-VMStartPolicy | Sort-Object VirtualMachineName
                        if ($VMStartPolicy) {
                            #region VM Autostart Section
                            Section -Style Heading2 'VM Autostart' {
                                Paragraph "The following section details the VM autostart configuration for $($VMHost.ExtensionData.Name)."
                                BlankLine
                                $VMStartPolicies = foreach ($VMStartPol in $VMStartPolicy) {
                                    [PSCustomObject]@{
                                        'Virtual Machine' = $VMStartPol.VirtualMachineName
                                        'Autostart Enabled' = Switch ($VMStartPol.StartAction) {
                                            'PowerOn' { 'Yes' }
                                            'None' { 'No' }
                                            default { $VMStartPol.StartAction }
                                        }
                                        'Autostart Order' = Switch ($VMStartPol.StartOrder) {
                                            $null { 'Unset' }
                                            default { $VMStartPol.StartOrder }
                                        }
                                        'Shutdown Behavior' = Switch ($VMStartPol.StopAction) {
                                            'PowerOff' { 'Power Off' }
                                            'GuestShutdown' { 'Shutdown' }
                                            default { $VMStartPol.StopAction }
                                        }
                                        'Start Delay' = "$($VMStartPol.StartDelay) sec"
                                        'Stop Delay' = "$($VMStartPol.StopDelay) sec"
                                        'Wait for Heartbeat' = Switch ($VMStartPol.WaitForHeartbeat) {
                                            $true { 'Yes' }
                                            $false { 'No' }
                                        }
                                    }
                                }
                                $TableParams = @{
                                    Name = "VM Autostart Policy - $($VMHost.ExtensionData.Name)"
                                    ColumnWidths = 25, 12, 13, 15, 12, 12, 11
                                }
                                if ($Report.ShowTableCaptions) {
                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                }
                                $VMStartPolicies | Table @TableParams
                            }
                            #endregion VM Autostart Section
                        }
                    }
                    #endregion ESXi Host VM Autostart Information
                }
                #endregion ESXi Host Detailed Information
            }
            #endregion Hosts Section
        } # end if ($ESXi)

        # Disconnect ESXi Server
        $Null = Disconnect-VIServer -Server $ESXi -Confirm:$false -ErrorAction SilentlyContinue

        #region Variable cleanup
        Clear-Variable -Name ESXi
        #endregion Variable cleanup

    } # end foreach ($VIServer in $Target)
}