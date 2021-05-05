$final_local = "$env:userprofile\Desktop\ADData\GroupData\";

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

$ErrorActionPreference = 'SilentlyContinue'
Get-ADGroup -Filter {GroupCategory -eq 'Security'} | Where-Object{@(Get-ADGroupMember $_).Length -eq 0} | Select-Object name,distinguishedName | Select-Object | export-csv "$final_local\EmptySecGroups_$date.txt" -NoTypeInformation
Get-ADGroup -Filter {GroupCategory -eq 'Distribution'} | Where-Object{@(Get-ADGroupMember $_).Length -eq 0} | Select-Object name,distinguishedName | Select-Object | export-csv "$final_local\EmptyDistroGroups_$date.txt" -NoTypeInformation
