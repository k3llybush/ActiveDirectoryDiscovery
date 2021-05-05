Import-Module -name GroupPolicy

$final_local = "$env:userprofile\Desktop\ADData\GPOData\";

$date = get-date -format M.d.yyyy
$local = Get-Location;

if (!$local.Equals("C:\")) { Set-Location "C:\" }

if ((Test-Path $final_local) -eq 0) {
    New-item -Path $final_local -ItemType "directory"
    Set-Location $final_local;
}
elseif ((Test-Path $final_local) -eq 1) {
    Set-Location $final_local
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