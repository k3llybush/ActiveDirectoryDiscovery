$local = Get-Location;
$date = get-date -format M.d.yyyy 
$final_local = "$env:userprofile\Desktop\ADData\GPO Data\Disabled\$date";

Import-Module grouppolicy 

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

import-module grouppolicy

(Get-ADForest).domains | Foreach-Object { Get-GPO -All -Domain $_ | Where-Object { $_.GPOStatus -eq "AllSettingsDisabled" } } `
 | Select-Object -Property DisplayName | Out-File "$final_local\disabled-gpos.txt"