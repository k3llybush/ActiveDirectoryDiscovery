#Create the path on the users Desktop
$date = get-date -format M.d.yyyy 
$local = Get-Location;
$final_local = "$env:userprofile\Desktop\ADData\ComputerData\$date";

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

#############

Get-ADComputer -Filter * -Property * | Select-Object Name,OperatingSystem,OperatingSystemServicePack,OperatingSystemVersion `
    | Export-CSV "$final_local\All-Computer-Data-OS.csv" -NoTypeInformation -Encoding UTF8