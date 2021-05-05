$final_local = "$env:userprofile\Desktop\ADData\OUData\";

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

Get-ADOrganizationalUnit -Filter * -Properties * | 
Where-Object {-not ( Get-ADObject -Properties * -Filter * -SearchBase $_.Distinguishedname `
    -SearchScope OneLevel -ResultSetSize 1 )} | Select-Object name,distinguishedName,Created | `
    Select-Object | export-csv "$final_local\empty-OUs_$date.txt" -NoTypeInformation



