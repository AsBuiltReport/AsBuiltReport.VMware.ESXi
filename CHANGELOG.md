# :arrows_clockwise: AsBuiltReport.VMware.ESXi Changelog

## [[1.1.4](https://github.com/AsBuiltReport/AsBuiltReport.VMware.ESXi/releases/tag/v1.1.4)] - 2025-08-22

### Changed
- Update module manifest `RequiredModules` updated for AsBuiltReport.Core 1.4.3
- Update PowerCLI module requirements to VCF PowerCLI 9.0
- Update colour placeholders in `README.md`
- Update `Get-RequiredModule` script function
- Update bug and feature request templates
- Change table column widths for list tables to 40/60
- Add try/catch code blocks for improved error handling

### Fixed
- ESXi storage information not shown when VMHost InfoLevel set to 0
- Update VMHost PCI Devices reporting to fix issues with ESXi 8.x hosts (@orb71)
- Fix issue with license reporting

### Removed
- Remove VMware document style script

## [[1.1.3](https://github.com/AsBuiltReport/AsBuiltReport.VMware.ESXi/releases/tag/v1.1.3)] - 2022-04-21

### Added
- Added VMHost IPMI / BMC configuration information

## [[1.1.2](https://github.com/AsBuiltReport/AsBuiltReport.VMware.ESXi/releases/tag/v1.1.2)] - 2022-03-24

### Added
- Automated tweet release workflow

### Fixed
- Fix hostname in virtual switch report section
- Fix colour placeholders in `README.md`

## [[1.1.0](https://github.com/AsBuiltReport/AsBuiltReport.VMware.ESXi/releases/tag/v1.1.0)] - 2021-10-09

### Added
- PowerShell 7 compatibility
- PSScriptAnalyzer & Release GitHub Action workflows
- VMHost network adapter LLDP reporting
- NSX TCP/IP stacks for VMkernel Adpater reporting
- Include release and issue links in `CHANGELOG.md`

### Changed
- VMkernel Adapter reporting for enabled services

### Fixed
- Display issues with highlights in `README.md`

## [[1.0.0](https://github.com/AsBuiltReport/AsBuiltReport.VMware.ESXi/releases/tag/v1.1.0)] - 2020-04-02
### Added
- Initial release of VMware ESXi As Built Report

### Fixed
- Created new VMware ESXi As Built Report ([Fix #1](https://github.com/AsBuiltReport/AsBuiltReport.VMware.ESXi/issues/1))