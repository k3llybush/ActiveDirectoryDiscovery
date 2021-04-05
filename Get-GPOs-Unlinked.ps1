$local = Get-Location;
$date = get-date -format M.d.yyyy 
$final_local = "$env:userprofile\Desktop\ADData\GPO Data\UnLinked\$date";

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

function IsNotLinked($xmldata){ 
    If ($xmldata.GPO.LinksTo -eq $null) { 
        Return $true 
    } 
     
    Return $false 
} 
 
$unlinkedGPOs = @() 
 
Get-GPO -All | ForEach { $gpo = $_ ; $_ | Get-GPOReport -ReportType xml | ForEach { If(IsNotLinked([xml]$_)){$unlinkedGPOs += $gpo} }} 
 
If ($unlinkedGPOs.Count -eq 0) { 
    "No Unlinked GPO's Found" 
} 
Else{ 
    $unlinkedGPOs | Select DisplayName,ID | ft 
    $unlinkedGPOs | Select DisplayName | Out-File "$final_local\unlinked-gpos.txt"
}