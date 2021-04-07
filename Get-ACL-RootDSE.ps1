$final_local = "$env:userprofile\Desktop\ADData\ACL\";

$date = get-date -format M.d.yyyy
$local = Get-Location;

if (!$local.Equals("C:\")) {
    Set-Location "C:\";
    if ((Test-Path $final_local) -eq 0) {
        mkdir $final_local;
        Set-Location $final_local;
    }

    elseif ((Test-Path $final_local) -eq 1) {
        Set-Location $final_local;
        Write-Output $final_local;
    }
}

Import-Module ActiveDirectory

# This array will hold the report output.
$report = @()

# Hide the errors for a couple duplicate hash table keys.
#$schemaIDGUID = @{}
### NEED TO RECONCILE THE CONFLICTS ###
$ErrorActionPreference = 'SilentlyContinue'


$root = Get-ADRootDSE | Select-Object -ExpandProperty defaultnamingcontext
$report += Get-Acl -Path "AD:\$root" |
Select-Object -ExpandProperty Access

# Dump the raw report out to a CSV file for analysis in Excel.
$report | Export-Csv -Path "$final_local\Root_Permissions_$date.csv" -NoTypeInformation

