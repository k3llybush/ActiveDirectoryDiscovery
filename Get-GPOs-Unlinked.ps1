Import-Module grouppolicy

$final_local = "$env:userprofile\Desktop\ADData\GPO Data\UnLinked\";

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

function IsNotLinked($xmldata) { 
    If ($xmldata.GPO.LinksTo -eq $null) { 
        Return $true 
    } 
     
    Return $false 
} 
 
$unlinkedGPOs = @() 
 
Get-GPO -All | ForEach-Object { $gpo = $_ ; $_ | Get-GPOReport -ReportType xml | ForEach-Object { If (IsNotLinked([xml]$_)) { $unlinkedGPOs += $gpo } } } 
 
If ($unlinkedGPOs.Count -eq 0) { 
    "No Unlinked GPO's Found" 
} 
Else { 
    $unlinkedGPOs | Select-Object DisplayName, ID | Format-Table 
    $unlinkedGPOs | Select-Object DisplayName | Out-File "$final_local\unlinked-gpos_$date.txt"
}