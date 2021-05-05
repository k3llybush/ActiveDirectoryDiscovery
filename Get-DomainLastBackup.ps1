$final_local = "$env:userprofile\Desktop\ADData\DomainInfo\$date";

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

$Domain = (Get-ADDomain).DNSRoot
$Regex = '\d\d\d\d-\d\d-\d\d'
$LastBackup = (repadmin /showbackup $Domain | Select-String $Regex | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value } )[0]

$BackupResult = "Last Active Directory backup occurred on $LastBackup"
$BackupResult | out-file $final_local\backup_$date.txt