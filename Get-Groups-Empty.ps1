$local = Get-Location;
$final_local = "$env:userprofile\Desktop\ADData\GroupData\";

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
$ErrorActionPreference = 'SilentlyContinue'
Get-ADGroup -Filter {GroupCategory -eq 'Security'} | ?{@(Get-ADGroupMember $_).Length -eq 0} | select name,distinguishedName | Select-Object | export-csv "$final_local\EmptySecGroups.txt"
Get-ADGroup -Filter {GroupCategory -eq 'Distribution'} | ?{@(Get-ADGroupMember $_).Length -eq 0} | select name,distinguishedName | Select-Object | export-csv "$final_local\EmptyDistroGroups.txt"
