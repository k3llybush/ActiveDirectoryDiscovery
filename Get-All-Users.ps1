PROCESS { #This is where the script executes  
    $final_local = "$env:userprofile\Desktop\ADData\AllADUsers\" 

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
    
    $csvreportfile = $final_local + "\ALLADUsers_$date.csv" 
    
    #import the ActiveDirectory Module 
    Import-Module ActiveDirectory 
     
    #Perform AD search. The quotes "" used in $SearchLoc is essential 
    #Without it, Export-ADUsers returuned error 
    Get-ADUser -Properties * -Filter * |  
    Select-Object @{Label = "First Name"; Expression = { $_.GivenName } },  
    @{Label = "Last Name"; Expression = { $_.Surname } }, 
    @{Label = "Display Name"; Expression = { $_.DisplayName } }, 
    @{Label = "Logon Name"; Expression = { $_.sAMAccountName } },
    @{Label = "Logon Script"; Expression = { $_.scriptpath } },
    @{Label = "Full address"; Expression = { $_.StreetAddress } }, 
    @{Label = "City"; Expression = { $_.City } }, 
    @{Label = "State"; Expression = { $_.st } }, 
    @{Label = "Post Code"; Expression = { $_.PostalCode } }, 
    @{Label = "Country/Region"; Expression = { if (($_.Country -eq 'GB')  ) { 'United Kingdom' } Else { '' } } }, 
    @{Label = "Job Title"; Expression = { $_.Title } }, 
    @{Label = "Company"; Expression = { $_.Company } }, 
    @{Label = "Description"; Expression = { $_.Description } }, 
    @{Label = "Department"; Expression = { $_.Department } }, 
    @{Label = "Office"; Expression = { $_.OfficeName } }, 
    @{Label = "Phone"; Expression = { $_.telephoneNumber } }, 
    @{Label = "Email"; Expression = { $_.Mail } }, 
    @{Label = "Account Status"; Expression = { if (($_.Enabled -eq 'TRUE')  ) { 'Enabled' } Else { 'Disabled' } } }, # the 'if statement# replaces $_.Enabled 
    @{Label = "Last LogOn Date"; Expression = { $_.lastlogondate } },  
    @{Label = "Last LogOn TimeStamp"; Expression = { $_.lastlogonTimeStamp } },
    @{Label = "sidHistory"; Expression = { $_.sidhistory } } |
                   
    #Export CSV report 
    Export-Csv -Path $csvreportfile -NoTypeInformation     
}