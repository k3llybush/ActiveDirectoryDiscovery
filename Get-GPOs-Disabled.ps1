Import-Module -name GroupPolicy

$final_local = "$env:userprofile\Desktop\ADData\GPOData\";

$date = get-date -format M.d.yyyy
$local = Get-Location;

if(!$local.Equals("C:\"))
{
    Set-Location "C:\";
    if((Test-Path $final_local) -eq 0)
    {
        mkdir $final_local;
        Set-Location $final_local;
    }

    ## if path already exists
    ## DB Connect
    elseif ((Test-Path $final_local) -eq 1)
    {
        Set-Location $final_local;
        Write-Output $final_local;
    }
}

(Get-ADForest).domains | Foreach-Object { Get-GPO -All -Domain $_ | Where-Object { $_.GPOStatus -eq "AllSettingsDisabled" } } `
 | Select-Object -Property DisplayName | Out-File "$final_local\gpos-disabled_$date.txt"