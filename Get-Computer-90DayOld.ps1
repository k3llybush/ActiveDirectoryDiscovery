$local = Get-Location;
$date = get-date -format M.d.yyyy 
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

# Gets time stamps for all computers in the domain that have NOT logged in since after specified date 
import-module activedirectory  
$DaysInactive = 90  
$time = (Get-Date).Adddays(-($DaysInactive)) 
  
# Get all AD computers with lastLogonTimestamp less than our time 
Get-ADComputer -Filter {LastLogonTimeStamp -lt $time} -Properties LastLogonTimeStamp | 
  
# Output hostname and lastLogonTimestamp into CSV 
select-object Name,@{Name="Stamp"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} `
    | export-csv "$final_local\OLD_Computer-90.csv" -notypeinformation

