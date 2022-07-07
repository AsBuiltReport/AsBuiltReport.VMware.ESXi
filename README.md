<p align="center">
    <a href="https://www.asbuiltreport.com/" alt="AsBuiltReport"></a>
            <img src='https://raw.githubusercontent.com/AsBuiltReport/AsBuiltReport/master/AsBuiltReport.png' width="8%" height="8%" /></a>
</p>
<p align="center">
    <a href="https://www.powershellgallery.com/packages/AsBuiltReport.VMware.ESXi/" alt="PowerShell Gallery Version">
        <img src="https://img.shields.io/powershellgallery/v/AsBuiltReport.VMware.ESXi.svg" /></a>
    <a href="https://www.powershellgallery.com/packages/AsBuiltReport.VMware.ESXi/" alt="PS Gallery Downloads">
        <img src="https://img.shields.io/powershellgallery/dt/AsBuiltReport.VMware.ESXi.svg" /></a>
    <a href="https://www.powershellgallery.com/packages/AsBuiltReport.VMware.ESXi/" alt="PS Platform">
        <img src="https://img.shields.io/powershellgallery/p/AsBuiltReport.VMware.ESXi.svg" /></a>
</p>
<p align="center">
    <a href="https://github.com/AsBuiltReport/AsBuiltReport.VMware.ESXi/graphs/commit-activity" alt="GitHub Last Commit">
        <img src="https://img.shields.io/github/last-commit/AsBuiltReport/AsBuiltReport.VMware.ESXi/master.svg" /></a>
    <a href="https://raw.githubusercontent.com/AsBuiltReport/AsBuiltReport.VMware.ESXi/master/LICENSE" alt="GitHub License">
        <img src="https://img.shields.io/github/license/AsBuiltReport/AsBuiltReport.VMware.ESXi.svg" /></a>
    <a href="https://github.com/AsBuiltReport/AsBuiltReport.VMware.ESXi/graphs/contributors" alt="GitHub Contributors">
        <img src="https://img.shields.io/github/contributors/AsBuiltReport/AsBuiltReport.VMware.ESXi.svg"/></a>
</p>
<p align="center">
    <a href="https://twitter.com/AsBuiltReport" alt="Twitter">
            <img src="https://img.shields.io/twitter/follow/AsBuiltReport.svg?style=social"/></a>
</p>

<p align="center">
    <a href='https://ko-fi.com/B0B7DDGZ7' target='_blank'><img height='36' style='border:0px;height:36px;' src='https://cdn.ko-fi.com/cdn/kofi1.png?v=3' border='0' alt='Buy Me a Coffee at ko-fi.com' /></a>
</p>

# VMware ESXi As Built Report

VMware ESXi As Built Report is a PowerShell module which works in conjunction with [AsBuiltReport.Core](https://github.com/AsBuiltReport/AsBuiltReport.Core).

[AsBuiltReport](https://github.com/AsBuiltReport/AsBuiltReport) is an open-sourced community project which utilises PowerShell to produce as-built documentation in multiple document formats for multiple vendors and technologies.

The VMware ESXi As Built Report module is used to generate as built documentation for standalone VMware ESXi servers.

Please refer to the [VMware vSphere AsBuiltReport](https://github.com/AsBuiltReport/AsBuiltReport.VMware.vSphere) for reporting of VMware vSphere / vCenter Server environments.

Please refer to the AsBuiltReport [website](https://www.asbuiltreport.com) for more detailed information about this project.

# :beginner: Getting Started
Below are the instructions on how to install, configure and generate a VMware ESXi As Built report.

## :floppy_disk: Supported Versions

### VMware ESXi
The VMware ESXi As Built Report supports the following ESXi versions;
- ESXi 6.5
- ESXi 6.7
- ESXi 7.0

#### End of Support
The following VMware ESXi versions are no longer being tested and/or supported;
- ESXi 5.5
- ESXi 6.0

### PowerShell
This report is compatible with the following PowerShell versions;

| Windows PowerShell 5.1 |     PowerShell 7    |
|:----------------------:|:--------------------:|
|   :white_check_mark:   | :white_check_mark: |

## :wrench: System Requirements
PowerShell 5.1 or PowerShell 7, and the following PowerShell modules are required for generating a VMware ESXi As Built report.

- [VMware PowerCLI Module](https://www.powershellgallery.com/packages/VMware.PowerCLI/)
- [AsBuiltReport.VMware.ESXi Module](https://www.powershellgallery.com/packages/AsBuiltReport.VMware.ESXi/)

### Linux & macOS
* .NET Core is required for cover page image support on Linux and macOS operating systems.
    * [Installing .NET Core for macOS](https://docs.microsoft.com/en-us/dotnet/core/install/macos)
    * [Installing .NET Core for Linux](https://docs.microsoft.com/en-us/dotnet/core/install/linux)

❗ If you are unable to install .NET Core, you must set `ShowCoverPageImage` to `False` in the report JSON configuration file.

### :closed_lock_with_key: Required Privileges
A user with root privileges on the ESXi host is required to generate a VMware ESXi As Built Report.

## :package: Module Installation

Open a PowerShell terminal window and install each of the required modules.

:warning: VMware PowerCLI 12.3 or higher is required. Please ensure older PowerCLI versions have been uninstalled.

```powershell
install-module VMware.PowerCLI -MinimumVersion 12.3 -AllowClobber
install-module AsBuiltReport.VMware.ESXi
```

## :pencil2:Configuration

The ESXi As Built Report utilises a JSON file to allow configuration of report information, options, detail and healthchecks.

An ESXi report configuration file can be generated by executing the following command;
```powershell
New-AsBuiltReportConfig -Report VMware.ESXi -Path <User specified folder> -Name <Optional>
```

Executing this command will copy the default ESXi report JSON configuration to a user specified folder.

All report settings can then be configured via the JSON file.

The following provides information of how to configure each schema within the report's JSON file.

### Report
The **Report** schema provides configuration of the ESXi report information

| Sub-Schema         | Setting      | Default                     | Description                                                   |
|--------------------|------------- | --------------------------- | --------------------------------------------------------------|
| Name               | User defined | VMware ESXi As Built Report | The name of the As Built Report                               |
| Version            | User defined | 1.0                         | The report version                                            |
| Status             | User defined | Released                    | The report release status                                     |
| ShowCoverPageImage | true / false | true                        | Toggle to enable/disable the display of the cover page image  |
| ShowHeaderFooter   | true / false | true                        | Toggle to enable/disable document headers & footers           |
| ShowTableCaptions  | true / false | true                        | Toggle to enable/disable table captions/numbering             |

### Options
The **Options** schema allows certain options within the report to be toggled on or off

| Sub-Schema         | Setting      | Default | Description                                                                                                                                                                              |
|--------------------|--------------|---------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| ShowLicenseKeys    | true / false | false   | Toggle to mask/unmask ESXi license keys<br><br> **Masked License Key**<br>\*\*\*\*\*-\*\*\*\*\*-\*\*\*\*\*-56YDM-AS12K<br><br> **Unmasked License Key**<br>AKLU4-PFG8M-W2D8J-56YDM-AS12K |
| ShowVMSnapshots    | true / false | true    | Toggle to enable/disable reporting of VM snapshots                                                                                                                                       |

### InfoLevel
The **InfoLevel** schema allows configuration of each section of the report at a granular level. The following sections can be set

There are 6 levels (0-5) of detail granularity for each section as follows;

| Setting | InfoLevel         | Description                                                                                                                                |
|---------|-------------------|--------------------------------------------------------------------------------------------------------------------------------------------|
| 0       | Disabled          | Does not collect or display any information                                                                                                |
| 1       | Enabled / Summary | Provides summarised information for a collection of objects                                                                                |
| 2       | Adv Summary       | Provides condensed, detailed information for a collection of objects                                                                       |
| 3       | Detailed          | Provides detailed information for individual objects                                                                                       |
| 4       | Adv Detailed      | Provides detailed information for individual objects, as well as information for associated objects (Hosts, Clusters, Datastores, VMs etc) |
| 5       | Comprehensive     | Provides comprehensive information for individual objects, such as advanced configuration settings                                         |

The table below outlines the default and maximum **InfoLevel** settings for each section.

| Sub-Schema | Default Setting | Maximum Setting |
|------------|:---------------:|:---------------:|
| VMHost     |        3        |        5        |
| Network    |        3        |        4        |
| Storage    |        3        |        4        |
| VM         |        3        |        4        |

### Healthcheck
The **Healthcheck** schema is used to toggle health checks on or off.

#### VMHost
The **VMHost** schema is used to configure health checks for VMHosts.

| Sub-Schema      | Setting      | Default | Description                                                                                                              | Highlight                                                                                                                                                                                       |
|-----------------|--------------|---------|--------------------------------------------------------------------------------------------------------------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| ConnectionState | true / false | true    | Checks VMHosts connection state                                                                                          | ![Warning](https://via.placeholder.com/15/FFE860/FFE860.png) Maintenance<br>  ![Critical](https://via.placeholder.com/15/FFB38F/FFB38F.png)  Disconnected                                               |
| HyperThreading  | true / false | true    | Highlights VMHosts which have HyperThreading disabled                                                                    | ![Warning](https://via.placeholder.com/15/FFE860/FFE860.png) HyperThreading disabled<br>                                                                                                            |
| ScratchLocation | true / false | true    | Highlights VMHosts which are configured with the default scratch location                                                | ![Warning](https://via.placeholder.com/15/FFE860/FFE860.png) Scratch location is /tmp/scratch                                                                                                       |
| IPv6            | true / false | true    | Highlights VMHosts which do not have IPv6 enabled                                                                        | ![Warning](https://via.placeholder.com/15/FFE860/FFE860.png) IPv6 disabled                                                                                                                          |
| UpTimeDays      | true / false | true    | Highlights VMHosts with uptime days greater than 9 months                                                                | ![Warning](https://via.placeholder.com/15/FFE860/FFE860.png) 9 - 12 months<br> ![Critical](https://via.placeholder.com/15/FFB38F/FFB38F.png)  >12 months                                                |
| Licensing       | true / false | true    | Highlights VMHosts which are using production evaluation licenses                                                        | ![Warning](https://via.placeholder.com/15/FFE860/FFE860.png) Product evaluation license in use                                                                                                      |
| SSH             | true / false | true    | Highlights if the SSH service is enabled                                                                                 | ![Warning](https://via.placeholder.com/15/FFE860/FFE860.png) TSM / TSM-SSH service enabled                                                                                                          |
| ESXiShell       | true / false | true    | Highlights if the ESXi Shell service is enabled                                                                          | ![Warning](https://via.placeholder.com/15/FFE860/FFE860.png) TSM / TSM-EsxiShell service enabled                                                                                                    |
| NTP             | true / false | true    | Highlights if the NTP service has stopped or is disabled on a VMHost                                                     | ![Critical](https://via.placeholder.com/15/FFB38F/FFB38F.png)  NTP service stopped / disabled                                                                                                       |
| StorageAdapter  | true / false | true    | Highlights storage adapters which are not 'Online'                                                                       | ![Warning](https://via.placeholder.com/15/FFE860/FFE860.png) Storage adapter status is 'Unknown'<br> ![Critical](https://via.placeholder.com/15/FFB38F/FFB38F.png)  Storage adapter status is 'Offline' |
| NetworkAdapter  | true / false | true    | Highlights physical network adapters which are not 'Connected'<br> Highlights physical network adapters which are 'Down' | ![Critical](https://via.placeholder.com/15/FFB38F/FFB38F.png)  Network adapter is 'Disconnected'<br> ![Critical](https://via.placeholder.com/15/FFB38F/FFB38F.png)  Network adapter is 'Down'           |
| LockdownMode    | true / false | true    | Highlights VMHosts which do not have Lockdown mode enabled                                                               | ![Warning](https://via.placeholder.com/15/FFE860/FFE860.png) Lockdown Mode disabled<br>                                                                                                             |

#### Datastore
The **Datastore** schema is used to configure health checks for Datastores.

| Sub-Schema          | Setting      | Default | Description                                                      | Highlight                                                                                                                                              |
|---------------------|--------------|---------|------------------------------------------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------|
| CapacityUtilization | true / false | true    | Highlights datastores with storage capacity utilization over 75% | ![Warning](https://via.placeholder.com/15/FFE860/FFE860.png) 75 - 90% utilized<br> ![Critical](https://via.placeholder.com/15/FFB38F/FFB38F.png) >90% utilized |

#### VM
The **VM** schema is used to configure health checks for virtual machines.

| Sub-Schema           | Setting      | Default | Description                                                                                          | Highlight                                                                                                                                                                                                           |
|----------------------|--------------|---------|------------------------------------------------------------------------------------------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| PowerState           | true / false | true    | Highlights VMs which are powered off                                                                 | ![Warning](https://via.placeholder.com/15/FFE860/FFE860.png) VM is powered off                                                                                                                                          |
| ConnectionState      | true / false | true    | Highlights VMs which are orphaned or inaccessible                                                    | ![Critical](https://via.placeholder.com/15/FFB38F/FFB38F.png) VM is orphaned or inaccessible                                                                                                                            |
| CpuHotAdd            | true / false | true    | Highlights virtual machines which have CPU Hot Add enabled                                           | ![Warning](https://via.placeholder.com/15/FFE860/FFE860.png) CPU Hot Add enabled                                                                                                                                        |
| CpuHotRemove         | true / false | true    | Highlights virtual machines which have CPU Hot Remove enabled                                        | ![Warning](https://via.placeholder.com/15/FFE860/FFE860.png) CPU Hot Remove enabled                                                                                                                                     |
| MemoryHotAdd         | true / false | true    | Highlights VMs which have Memory Hot Add enabled                                                     | ![Warning](https://via.placeholder.com/15/FFE860/FFE860.png) Memory Hot Add enabled                                                                                                                                     |
| ChangeBlockTracking  | true / false | true    | Highlights VMs which do not have Change Block Tracking enabled                                       | ![Warning](https://via.placeholder.com/15/FFE860/FFE860.png) Change Block Tracking disabled                                                                                                                             |
| SpbmPolicyCompliance | true / false | true    | Highlights VMs which do not comply with storage based policies                                       | ![Warning](https://via.placeholder.com/15/FFE860/FFE860.png) VM storage based policy compliance is unknown<br> ![Critical](https://via.placeholder.com/15/FFB38F/FFB38F.png) VM does not comply with storage based policies |
| VMToolsStatus        | true / false | true    | Highlights Virtual Machines which do not have VM Tools installed, are out of date or are not running | ![Warning](https://via.placeholder.com/15/FFE860/FFE860.png) VM Tools not installed, out of date or not running                                                                                                         |
| VMSnapshots          | true / false | true    | Highlights Virtual Machines which have snapshots older than 7 days                                   | ![Warning](https://via.placeholder.com/15/FFE860/FFE860.png) VM Snapshot age >= 7 days<br> ![Critical](https://via.placeholder.com/15/FFB38F/FFB38F.png) VM Snapshot age >= 14 days                                         |

## :computer: Examples

```powershell
# Generate an ESXi As Built Report for ESXi server 'esxi-01.corp.local' using specified credentials. Export report to HTML & DOCX formats. Use default report style. Append timestamp to report filename. Save reports to 'C:\Users\Tim\Documents'
PS C:\> New-AsBuiltReport -Report VMware.ESXi -Target 'esxi-01.corp.local' -Username 'root' -Password 'VMware1!' -Format Html,Word -OutputFolderPath 'C:\Users\Tim\Documents' -Timestamp

# Generate an ESXi As Built Report for ESXi server 'esxi-01.corp.local' using specified credentials and report configuration file. Export report to Text, HTML & DOCX formats. Use default report style. Save reports to 'C:\Users\Tim\Documents'. Display verbose messages to the console.
PS C:\> New-AsBuiltReport -Report VMware.ESXi -Target 'esxi-01.corp.local' -Username 'root' -Password 'VMware1!' -Format Text,Html,Word -OutputFolderPath 'C:\Users\Tim\Documents' -ReportConfigFilePath 'C:\Users\Tim\AsBuiltReport\AsBuiltReport.VMware.ESXi.json' -Verbose

# Generate an ESXi As Built Report for ESXi server 'esxi-01.corp.local' using stored credentials. Export report to HTML & Text formats. Use default report style. Highlight environment issues within the report. Save reports to 'C:\Users\Tim\Documents'.
PS C:\> $Creds = Get-Credential
PS C:\> New-AsBuiltReport -Report VMware.ESXi -Target 'esxi-01.corp.local' -Credential $Creds -Format Html,Text -OutputFolderPath 'C:\Users\Tim\Documents' -EnableHealthCheck

# Generate a single ESXi As Built Report for ESXi servers 'esxi-01.corp.local' and 'esxi-02.corp.local' using specified credentials. Report exports to Word format by default. Apply custom style to the report. Reports are saved to the user profile folder by default.
PS C:\> New-AsBuiltReport -Report VMware.ESXi -Target 'esxi-01.corp.local','esxi-02.corp.local' -Username 'root' -Password 'VMware1!' -StylePath 'C:\Scripts\Styles\MyCustomStyle.ps1'

# Generate an ESXi As Built Report for ESXi server 'esxi-01.corp.local' using specified credentials. Export report to HTML & DOCX formats. Use default report style. Reports are saved to the user profile folder by default. Attach and send reports via e-mail.
PS C:\> New-AsBuiltReport -Report VMware.ESXi -Target 'esxi-01.corp.local' -Username 'root' -Password 'VMware1!' -Format Html,Word -OutputFolderPath 'C:\Users\Tim\Documents' -SendEmail
```