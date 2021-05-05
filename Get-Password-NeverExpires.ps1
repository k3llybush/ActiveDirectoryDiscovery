$final_local = "$env:userprofile\Desktop\ADData\AllADUsers\";

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

Search-ADAccount -PasswordNeverExpires | Export-Csv "$final_local\neverexpires.csv_$date" -NoTypeInformation