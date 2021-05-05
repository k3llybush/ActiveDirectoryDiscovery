function Get-GPOMappings {

Clear-Host

Add-Type -AssemblyName PresentationCore,PresentationFramework

$DomainObject       = Get-ADDomain
$DOmainName         = $DomainObject.DNSRoot
$PDCEmulator        = $DomainObject.PDCEmulator
[array]$GPOs        = @(Get-GPO -All -Domain $DOmainName)
try
{
    $SYSVOL         = Get-ChildItem -Path "\\$PDCEmulator\SYSVOL\$DOmainName\Policies" -ErrorAction Stop
}
catch
{
    $ButtonType         = [System.Windows.MessageBoxButton]::Ok
    $MessageboxTitle    = “Error contacting PDC Emulator”
    $MessageIcon        = [System.Windows.MessageBoxImage]::Error
    $Messageboxbody     = @“
Error connecting to 'SYSVOL' folder of PDC Emulator at ($PDCEmulator). Fix this problem before going further.
”@
    [System.Windows.MessageBox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
    break;
}
[array]$GPOIds      = ($GPOs).ID
[array]$SYSVOLIds   = ((($SYSVOL).Name) -replace '{','') -replace '}',''
[array]$ResultIds   = $null
[array]$Report      = $null
[array]$Phantoms    = $null
$CountSysvol        = $SYSVOL.Count
$CountGPOs          = $GPOs.Count
$ButtonType         = [System.Windows.MessageBoxButton]::Ok
$MessageboxTitle    = “Summary Information for ($DOmainName)”


if($CountSysvol -eq $CountGPOs)
{
    foreach($Item in $SYSVOL)
    {
        $ID = (($Item.Name) -replace '{','') -replace '}',''

        try
        {
            if(Get-GPO -Guid $ID -Domain $DOmainName -ErrorAction Stop)
            {
                $GPO    = Get-GPO -Guid $ID -Domain $DOmainName
                $Name   = $GPO.DisplayName
                $Status = "OK!"
            }
        }

        catch
        {
            $Status    = "GPO Missing"
            $Name      = "N/A"
            $Phantoms += $ID
        }
    

        $Obj = New-Object -TypeName PSObject -Property @{
                "Name"   = $Name
                "Status" = $Status
                "GUID"   = $ID
                                                        }
        $Report += $Obj
    }

$Messageboxbody  = @“
Below information taken from ($PDCEmulator) which is PDC Emulator of ($DOmainName):

    SysVol: $CountSysvol
    GPO   : $CountGPOs

It seems you have no inconsistency issues between SYSVOL and GPOs on your PDC Emulator.
”@ 

$MessageIcon     = [System.Windows.MessageBoxImage]::Information

}

elseif($CountSysvol -gt $CountGPOs)
{
    foreach($Item in $SYSVOL)
        {
            $ID = (($Item.Name) -replace '{','') -replace '}',''

            try
            {
                if(Get-GPO -Guid $ID -Domain $DOmainName -ErrorAction Stop)
                {
                    $GPO    = Get-GPO -Guid $ID -Domain $DOmainName
                    $Name   = $GPO.DisplayName
                    $Status = "OK!"
                }
            }

            catch
            {
                $Status    = "Phantom"
                $Name      = "N/A"
                $Phantoms += $ID
            }
    

            $Obj = New-Object -TypeName PSObject -Property @{
                    "Name"   = $Name
                    "Status" = $Status
                    "GUID"   = $ID
                                                            }
            $Report += $Obj
        }

    $CountPhantoms = $Phantoms.count
    $Messageboxbody  = @“
Below information taken from ($PDCEmulator) which is PDC Emulator of ($DOmainName):

    SysVol: $CountSysvol
    GPO   : $CountGPOs

You probably need to carefuly remove items marked with 'Phantom' from you SYSVOL. There are about $CountPhantoms phantom folders in SYSVOL directory of PDC Emulator. See the list by clicking 'OK'.
”@
    $MessageIcon     = [System.Windows.MessageBoxImage]::Warning


}

elseif($CountSysvol -lt $CountGPOs)
{
    foreach($GPO in $GPOs)
    {
        $Name = $GPO.DisplayName
        $ID   = $GPO.Id
        
        if($ID -in $SYSVOLIds)
        {
            $Status = "OK!"
        }

        else
        {
            $Status    = "Folder Missing"
            $ResultIds = $ResultIds + $ID
        }

        $Obj = New-Object -TypeName PSObject -Property @{
                    "Name"   = $Name
                    "Status" = $Status
                    "GUID"   = $ID
                                                            }
        $Report += $Obj
    }

    $Messageboxbody  = @“
Below information taken from ($PDCEmulator) which is PDC Emulator of ($DOmainName):

    SysVol: $CountSysvol
    GPO   : $CountGPOs

It seems there are problems with some GPOs. There is a high chance that you have GPOs where there is no associated folder for them in SYSVOL directory of PDC.


Below policies are missing from SYSVOL of PDC Emulator. Possible solutions are restoring this GPOS from backup and import it on PDC or, change the PDC to the DC who has these missing GPOs in their SYSVOL:
    
    $ResultIds
”@ 

    $MessageIcon     = [System.Windows.MessageBoxImage]::Error

}


[System.Windows.MessageBox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon) | Out-Null
$Report | Out-GridView -Title "SYSVOL & GPO Reports of PDC ($PDCEmulator)"
}
