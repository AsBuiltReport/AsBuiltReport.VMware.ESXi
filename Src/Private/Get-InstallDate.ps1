function Get-InstallDate {
    $thisUUID = $esxcli.system.uuid.get.Invoke()
    $decDate = [Convert]::ToInt32($thisUUID.Split("-")[0], 16)
    $installDate = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($decDate))
    [PSCustomObject][Ordered]@{
        Name = $VMHost.ExtensionData.Name
        InstallDate = $installDate
    }
}