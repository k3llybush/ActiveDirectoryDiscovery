$final_local = "$env:userprofile\Desktop\ADData\GroupData\";

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
$ErrorActionPreference = 'SilentlyContinue'
Get-ADGroup -Filter {GroupCategory -eq 'Security'} | Where-Object{@(Get-ADGroupMember $_).Length -eq 0} | Select-Object name,distinguishedName | Select-Object | export-csv "$final_local\EmptySecGroups_$date.txt" -NoTypeInformation
Get-ADGroup -Filter {GroupCategory -eq 'Distribution'} | Where-Object{@(Get-ADGroupMember $_).Length -eq 0} | Select-Object name,distinguishedName | Select-Object | export-csv "$final_local\EmptyDistroGroups_$date.txt" -NoTypeInformation
