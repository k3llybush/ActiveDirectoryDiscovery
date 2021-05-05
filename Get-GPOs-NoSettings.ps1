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

function HasNoSettings {
    $cExtNodes = $xmldata.DocumentElement.SelectNodes($cQueryString, $XmlNameSpaceMgr)
  
    foreach ($cExtNode in $cExtNodes) {
        If ($cExtNode.HasChildNodes) {
            Return $false
        }
    }
    
    $uExtNodes = $xmldata.DocumentElement.SelectNodes($uQueryString, $XmlNameSpaceMgr)
    
    foreach ($uExtNode in $uExtNodes) {
        If ($uExtNode.HasChildNodes) {
            Return $false
        }
    }
    
    Return $true
}

function configNamespace {
    $script:xmlNameSpaceMgr = New-Object System.Xml.XmlNamespaceManager($xmldata.NameTable)

    $xmlNameSpaceMgr.AddNamespace("", $xmlnsGpSettings)
    $xmlNameSpaceMgr.AddNamespace("gp", $xmlnsGpSettings)
    $xmlNameSpaceMgr.AddNamespace("xsi", $xmlnsSchemaInstance)
    $xmlNameSpaceMgr.AddNamespace("xsd", $xmlnsSchema)
}

$noSettingsGPOs = @()

$xmlnsGpSettings = "http://www.microsoft.com/GroupPolicy/Settings"
$xmlnsSchemaInstance = "http://www.w3.org/2001/XMLSchema-instance"
$xmlnsSchema = "http://www.w3.org/2001/XMLSchema"

$cQueryString = "gp:Computer/gp:ExtensionData/gp:Extension"
$uQueryString = "gp:User/gp:ExtensionData/gp:Extension"

Get-GPO -All | ForEach-Object { $gpo = $_ ; $_ | Get-GPOReport -ReportType xml | ForEach-Object { $xmldata = [xml]$_ ; configNamespace ; If (HasNoSettings) { $noSettingsGPOs += $gpo } } }

If ($noSettingsGPOs.Count -eq 0) {
    "No GPO's Without Settings Were Found"
}
Else {
    $noSettingsGPOs | Select-Object DisplayName, ID | Format-Table
    $noSettingsGPOs | Select-Object DisplayName, ID | Export-Csv "$final_local\gpos-no-settings_$date.csv" -NoTypeInformation
}