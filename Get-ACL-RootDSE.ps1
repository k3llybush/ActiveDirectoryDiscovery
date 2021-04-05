Clear-Variable "final_local"

$local = Get-Location;
$final_local = "$env:userprofile\Desktop\ADData\ACL\";

if(!$local.Equals("C:\"))
{
    cd "C:\";
    if((Test-Path $final_local) -eq 0)
    {
        mkdir $final_local;
        cd $final_local;
    }

    ## if path already exists
    ## DB Connect
    elseif ((Test-Path $final_local) -eq 1)
    {
        cd $final_local;
        echo $final_local;
    }
}

Import-Module ActiveDirectory

# This array will hold the report output.
$report = @()

# Hide the errors for a couple duplicate hash table keys.
$schemaIDGUID = @{}
### NEED TO RECONCILE THE CONFLICTS ###
$ErrorActionPreference = 'SilentlyContinue'


$root = Get-ADRootDSE | Select -ExpandProperty defaultnamingcontext
$report += Get-Acl -Path "AD:\$root" |
  Select-Object -ExpandProperty Access

# Dump the raw report out to a CSV file for analysis in Excel.
$report | Export-Csv -Path "$final_local\Root_Permissions.csv" -NoTypeInformation

