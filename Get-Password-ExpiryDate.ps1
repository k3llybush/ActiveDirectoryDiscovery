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

Get-ADUser -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} –Properties “DisplayName”, “msDS-UserPasswordExpiryTimeComputed” |

Select-Object -Property “Displayname”,@{Name=“ExpiryDate”;Expression={[datetime]::FromFileTime($_.“msDS-UserPasswordExpiryTimeComputed”)}} | Export-Csv "$final_local\password-expires-date_$date.csv" -NoTypeInformation
