$local = Get-Location;
$final_local = "$env:userprofile\Desktop\ADData\DomainControllers\";

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

$domain = [System.Directoryservices.Activedirectory.Domain]::GetCurrentDomain()
$domain | ForEach-Object {$_.DomainControllers} | 
ForEach-Object {
  $hostEntry= [System.Net.Dns]::GetHostByName($_.Name)
  New-Object -TypeName PSObject -Property @{
      Name = $_.Name
      IPAddress = $hostEntry.AddressList[0].IPAddressToString
     }
} | Export-CSV "$env:userprofile\Desktop\ADData\DomainControllers\DCIP.csv" -NoTypeInformation -Encoding UTF8