Import-Module -name GroupPolicy

#$date = get-date -uformat "%Y_%m_%d_%I%M%p"
$final_local = "$env:userprofile\Desktop\ADData\GPO Data\GPO-InventoryReport\";
$local = Get-Location;
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


$CSVpath = "$final_local\GPO_Inventory.csv"

[array]$Report = @()

$GPOs = Get-GPO -all | Sort-Object Displayname

foreach ($GPO in $GPOs) 
{
    Write-Host "Processing GPO $($GPO.DisplayName)"
    $XMLReport = Get-GPOReport -GUID $($GPO.id) -ReportType xml
    $XML = [xml]$XMLReport
    
	$Types = @("User","Computer")
	
    Foreach ($Type in $Types)
	{
	#Write-Host "Processing $Type GPO $($GPO.DisplayName)"
        $ExtArray = $xml.gpo.$Type.ExtensionData | foreach-Object -process {$_.name}
        
        if ($Type -eq "User"){$UserExtEnabled = $xml.gpo.$type.Enabled}
        if ($Type -eq "Computer"){$ComputerExtEnabled = $xml.gpo.$type.Enabled}
                        
        $ExtCount = $ExtArray.count
        #write-host "Extension count is $ExtCount"
        	    
        if (($ExtCount -eq $Null) -or ($ExtCount -eq 0))
	    {
	        #write-host "$Type is False"
            if ($Type -eq "User"){$UserExtEmpty = "No Settings"}
            if ($Type -eq "Computer"){$ComputerExtEmpty = "No Settings"}
	    }
        
        Else
        {	
            #write-host "$Type is True"
            if ($Type -eq "User"){$UserExtEmpty = "Has Settings"}
            if ($Type -eq "Computer"){$ComputerExtEmpty = "Has Settings"}
	    }
    }
    #write-host "Building Report"
    #write-host "Computer EXT $ComputerExtEnabled"
    #write-host "User Ext $UserExtEnabled"
    $Report += New-Object PSObject -Property @{
	        'GPO Name' = $xml.gpo.name
            'User GPO Side Enabled' = $global:UserExtEnabled 
            'Computer GPO Side Enabled' = $global:ComputerExtEnabled 
	        'Has Computer Settings' = $ComputerExtEmpty
            'Has User Settings' = $UserExtEmpty
            'GPO Status' = $GPO.GpoStatus
            'Last Modified' = $GPO.ModificationTime
            'Created on' = $GPO.CreationTime	                
	        } | select-object 'GPO Name','User GPO Side Enabled','Has User Settings','Computer GPO Side Enabled','Has Computer Settings','GPO Status','Last Modified','Created on' | Export-CSv -NoTypeInformation -path $CSVpath -Append        
 
    	Clear-variable UserExtEmpty
        Clear-variable ComputerExtEmpty
        Clear-variable UserExtEnabled
        Clear-variable ComputerExtEnabled
	    Clear-Variable ExtArray
        Clear-Variable ExtCount
}
