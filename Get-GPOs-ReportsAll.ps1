Import-Module -name GroupPolicy

$final_local = "$env:userprofile\Desktop\ADData\GPOData\ReportsAll\$date";

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

Set-Location $final_local

$Domain = (Get-ADDomain).DNSRoot
$PDC = (Get-ADDomain).PDCEmulator

# Get the reports
Get-GPOReport -All -Domain $Domain -Server $PDC -ReportType HTML -Path "$final_local\!GPOReportsAll.HTML"
Get-GPO -All | % {$_.GenerateReport('html') | Out-File "$($_.DisplayName).html"}




