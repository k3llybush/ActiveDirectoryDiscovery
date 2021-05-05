$final_local = "$env:userprofile\Desktop\ADData\ComputerData\";

$date = get-date -format M.d.yyyy
$local = Get-Location;

if (!$local.Equals("C:\")) { Set-Location "C:\" }

if ((Test-Path $final_local) -eq 0) {
    New-item -Path $final_local -ItemType "directory"
    Set-Location $final_local;
}
elseif ((Test-Path $final_local) -eq 1) {
    Set-Location $final_local
}

Get-ADComputer -Filter * -Property * | Select-Object Name, OperatingSystem, OperatingSystemServicePack, OperatingSystemVersion `
| Export-CSV "$final_local\All-Computer-Data-OS_$date.csv" -NoTypeInformation -Encoding UTF8