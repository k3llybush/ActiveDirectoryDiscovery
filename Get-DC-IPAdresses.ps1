﻿$final_local = "$env:userprofile\Desktop\ADData\DomainControllers\";

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

$domain = [System.Directoryservices.Activedirectory.Domain]::GetCurrentDomain()
$domain | ForEach-Object { $_.DomainControllers } | 
ForEach-Object {
    $hostEntry = [System.Net.Dns]::GetHostByName($_.Name)
    New-Object -TypeName PSObject -Property @{
        Name      = $_.Name
        IPAddress = $hostEntry.AddressList[0].IPAddressToString
    }
} | Export-CSV "$env:userprofile\Desktop\ADData\DomainControllers\DCIP_$date.csv" -NoTypeInformation -Encoding UTF8