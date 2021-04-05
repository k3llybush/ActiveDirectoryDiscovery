$ErrorActionPreference = 'SilentlyContinue'

$date = get-date -format M.d.yyyy 
$local = Get-Location;
$final_local = "$env:userprofile\Desktop\ADData\DomainControllers";

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

## Check a Domain Controller for Journal Wrap error in the File Replication Service log

$domain = [System.Directoryservices.Activedirectory.Domain]::GetCurrentDomain()
$domain | ForEach-Object {$_.DomainControllers} |
 ForEach-Object {
 $DC = $_.NAME
 $DCHostName = $DC
 Write-Output “Checking the File Replication Service event log for Journal Wrap events on $DCHostName… `r”
## Grabs the most recent event in the File Replication Service event log.
## The last event is logged usually Journal Wrap
## Added Event 13503 to get if the FRS is stopped, would be the last entry if no 13568 is there
 $DCEvents = Get-EventLog -ComputerName $DCHostName -LogName “File Replication Service” -Newest 1
 $DCFRSEvents = $DCEvents | Select-Object TimeGenerated,EventID,Source,Message

ForEach ($Event in $DCFRSEvents)
 {
 IF ($Event.EventID -eq “13568”)
    {
        Write-Output “$DCHostName has logged a Journal Wrap in the event log." | Out-File C:\Users\adm_kwbush\Desktop\ADData\DomainControllers\JournalWrap.txt -append
    }
 }
  IF ($Event.EventID -eq “13503”)
    {
        Write-Output “$DCHostName has logged a FRS stop in the event log." | Out-File C:\Users\adm_kwbush\Desktop\ADData\DomainControllers\JournalWrap-Stop.txt -append
    }
 }