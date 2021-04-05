$local = Get-Location;

$date = get-date -format M.d.yyyy 

$final_local = "$env:userprofile\Desktop\ADData\GPO Data\Reports\$date";

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

Set-Location $final_local

$Domain = (Get-ADDomain).DNSRoot
$PDC = (Get-ADDomain).PDCEmulator

# Get the reports
Get-GPOReport -All -Domain $Domain -Server $PDC -ReportType HTML -Path "$final_local\!GPOReportsAll.HTML"
Get-GPO -All | % {$_.GenerateReport('html') | Out-File "$($_.DisplayName).html"}




