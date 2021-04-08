$final_local = "$env:userprofile\Desktop\ADData\DomainsAndTrust";

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

Import-Module ActiveDirectory
$strDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name
$file = $strDomain + "-domainsandtrusts_$date.txt"
$trusts = Get-ADObject -Filter { ObjectClass -eq "trustedDomain" } -Properties * | out-file $final_local\$file