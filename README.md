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

# AsBuiltReport.VMware.ESXi

# Getting Started
Below are the instructions on how to install, configure and generate a VMware ESXi As Built report.

## Supported ESXi Versions
The VMware ESXi As Built Report supports the following ESXi versions;
- ESXi 5.0
- ESXi 5.1
- ESXi 5.5
- ESXi 6.0
- ESXi 6.5
- ESXi 6.7

## Pre-requisites
The following PowerShell modules are required for generating a VMware ESXi As Built report.

Each of these modules can be easily downloaded and installed via the PowerShell Gallery 

- [VMware PowerCLI Module](https://www.powershellgallery.com/packages/VMware.PowerCLI/)
- [AsBuiltReport.VMware.ESXi Module](https://www.powershellgallery.com/packages/AsBuiltReport.VMware.ESXi/)

## Module Installation

Open a Windows PowerShell terminal window and install each of the required modules. 

**Note:** VMware PowerCLI 10.0 or higher required.

```powershell
install-module VMware.PowerCLI -MinimumVersion 10.0
install-module AsBuiltReport.VMware.ESXi
```

### Required Privileges