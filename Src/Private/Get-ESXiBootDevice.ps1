function Get-ESXiBootDevice {
    <#
    .NOTES
    ===========================================================================
        Created by:    William Lam
        Organization:  VMware
        Blog:          www.virtuallyghetto.com
        Twitter:       @lamw
    ===========================================================================
    #>
    $results = @()
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
            $bootVendor = $device.Vendor
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