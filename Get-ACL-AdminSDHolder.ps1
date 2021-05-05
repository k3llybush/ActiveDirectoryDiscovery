$final_local = "$env:userprofile\Desktop\ADData\ACL\";

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

Import-Module ActiveDirectory

# This array will hold the report output.
$report = @()

# Hide the errors for a couple duplicate hash table keys.
#$schemaIDGUID = @{}
### NEED TO RECONCILE THE CONFLICTS ###
$ErrorActionPreference = 'SilentlyContinue'

#ent.ad.ntrs.com/System/AdminSDHolder

$root = Get-ADRootDSE | Select-Object -ExpandProperty defaultnamingcontext

$rootplus = "CN=AdminSDHolder,CN=System," + $root

$report += Get-Acl -Path "AD:\$rootplus" | Select-Object -ExpandProperty Access

# Dump the raw report out to a CSV file for analysis in Excel.
$report | Export-Csv -Path "$final_local\Root_Permissions_$date.csv" -NoTypeInformation

