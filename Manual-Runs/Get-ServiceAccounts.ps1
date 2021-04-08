#—————————————————–# 
#Create the path on the users Desktop
$date = get-date -format M.d.yyyy 

$local = Get-Location;
$final_local = "$env:userprofile\Desktop\ADData\Services\";

if (!$local.Equals("C:\")) {
    Set-Location "C:\";
    if ((Test-Path $final_local) -eq 0) {
        mkdir $final_local;
        Set-Location $final_local;
    }

    elseif ((Test-Path $final_local) -eq 1) {
        Set-Location $final_local;
        Write-Output $final_local;
    }
}

#—————————————————–# 
#Imports the Active Directory PowerShell module 

Import-Module ActiveDirectory    

#—————————————————–#
# Gets all servers in the domain 

$servers = (Get-ADComputer -Filter { OperatingSystem -Like "Windows *Server*" }).name

#—————————————————–#
# Use below for test, comment out the line from running then uncomment one below
#$servers= "NHC0436"
#$servers= "MDT224"

#—————————————————–#
# For Each Server, find services running under the user specified in $Account 

ForEach ($server in $servers) { 

    if (-not (Test-Connection "$server" -Quiet -Count 1))
    {
        Write-Host "$($server): Unable to ping." -ForegroundColor Red
    }
    Else { 
        Write-Host "$($server): Is up." -ForegroundColor Green
        Try {
                        
            #—————————————————–#
            #
            # Pull services from the remote server and filter out LocalService, LocalSytem, NT Authority, NT Services
            #
            #—————————————————–#
                        
            $Services = Get-wmiobject win32_service -computername $server `
                -filter "NOT (startname LIKE 'LocalService%' OR startname LIKE 'LocalSystem%' OR startname LIKE 'NT Authority%' OR startname LIKE 'NT Service%' )" `
                #| Select Name,Displayname,Start*
        }
        Catch {
            if ($_.Exception.GetType().Name -eq "*") {
                Write-Host "$server error............." -ForegroundColor Yellow
            }
            Write-Host "$server error!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Cyan 
        }
        #—————————————————–#
        # List the services running in the powershell console 
        # If there are no services running, output this to the console. 

        If ($Services -ne $null) { 
            Write-Host $Services
            $Services | Export-Csv -Append "$final_local\services_$date.csv" -notypeinformation
        } 
        Elseif ($Services -eq $null) { 
            Write-Host "No Services found running on Server $server" 
        }
    }  
    #—————————————————–# 
    #End
}

#Clean-Memory

Function Clean-Memory {
    Get-Variable |
    Where-Object { $startupVariables -notcontains $_.Name } |
    ForEach-Object {
        try { Remove-Variable -Name "$($_.Name)" -Force -Scope "global" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue }
        catch { }
    }
}