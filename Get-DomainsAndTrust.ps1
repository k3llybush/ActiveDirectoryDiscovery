
$local = Get-Location;
$date = get-date -format M.d.yyyy 
$final_local = "$env:userprofile\Desktop\ADData\DomainsAndTrust";

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
$strDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name
$file = $strDomain + "-domainsandtrusts.txt"
$trusts = Get-ADObject -Filter { ObjectClass -eq "trustedDomain" } -Properties * | out-file $final_local\$file