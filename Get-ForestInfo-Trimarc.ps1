<#.
SCRIPTNAME: Get-ADForestInfo.ps1
AUTHOR: Sean Metcalf
AUTHOR EMAIL: sean@trimarcsecurity.com
COMPANY: Trimarc Security, LLC (Trimarc)
WEBSITE: https://www.TrimarcSecurity.com
COPYRIGHT: © 2018 - 2019 Trimarc Security, LLC (Trimarc)
This script and all elements within it are intellectual property owned by Trimarc Security, LLC (Trimarc) and are not to be used without permission.
No warranty or suitability of use is stated or implied.

Last Updated: 5/07/2019
Version 1.1.4

Pre-Requisites:
=======================================
   * Active Directory PowerShell Module
   * Group Policy PowerShell Module


#>

$final_local = "$env:userprofile\Desktop\ADData\DomainInfo\$date";

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

$ForestReportFile =  $final_local + "\ForestReport.log"
Start-Transcript -Path $ForestReportFile 

# Check for AD & GPO Modules
# If the Active Directory PowerShell module isn't installed, it can easily be installed on servers (Windows Server 2008 R2 and newer) by running the following as an Admin:

IF (!(Get-Command -module ActiveDirectory))
 { 
   Import-Module ServerManager
   Add-WindowsFeature RSAT-AD-PowerShell 
 }
ELSE
 { Write-Output "The Active Directory PowerShell Module is installed." }

IF (!(Get-Command -module GroupPolicy))
 { 
   Import-Module ServerManager
   Add-WindowsFeature –Name GPMC 
 }
ELSE
 { Write-Output "The GroupPolicy PowerShell Module is installed." }

######
## Set Windows OS Array
######
$WindowsServerOperatingSystemArray = @(
"Windows Server 2003",
"Windows Server 2008",
"Windows Server 2008 R2",
"Windows Server 2012",
"Windows Server 2012 R2",
"Windows Server 2016",
"Windows Server 2016 R2",
"Windows Server 2019"
)

#############################################
# Import Active Directory PowerShell Module #
#############################################
import-module ActiveDirectory
import-module GroupPolicy


#######################
## Get Forest Config ##
#######################
IF ($ForestDNSName)
    { $ADForestInfo = Get-ADForest $ForestDNSName }
  ELSE
    { $ADForestInfo = Get-ADForest }

# Set initial AD Forest Variables
$ADForestInfoName = $ADForestInfo.Name
$ForestRootDomain = $ADForestInfo.RootDomain
$ForestDNSName = $ADForestInfo.Name
$ForestDomainNetBIOSName = (Get-ADDomain $ADForestInfo.Name).NetBIOSName
$ForestFunctionalLevel = $ADForestInfo.ForestMode
$ForestDomains = $ADForestInfo.Domains
$ForestDomainCount = $ForestDomains.Count

[int]$TotalDomainsInForest = ($ADForestInfo.Domains).Count
[int]$TotalSitesInForest = ($ADForestInfo.Sites).Count

$TotalGlobalCatalogsInForest = ($ADForestInfo.GlobalCatalogs).Count

# Determine when the AD Forest was created
[string]$ForestDC = (Get-ADDomainController -discover -domain $ForestRootDomain).HostName
$ADInstatiationObject = Get-ADObject -SearchBase $ADForestInfo.PartitionsContainer -Server $ForestDC `
-LDAPFilter "(&(objectClass=crossRef)(systemFlags=3))" `
-Property dnsRoot,nETBIOSName,whenCreated | Sort-Object whenCreated 

$ForestCreationDate = $ADInstatiationObject | where {$_.dnsRoot -eq $ForestDNSName } | select whenCreated -expandproperty whenCreated


################################
### Domain Statistics ###
################################
$ForestDomainFunctionalLevels = $NULL
$ForestAdminCount1AccountUsers = @()
$ForestAdminCount1AccountUserCount = 0
$ForestUserCount = 0 
$ForestComputerRunningWindowsServerCount = 0
$ForestComputerRunningWindowsWorkstationCount = 0
$ForestDCCount = 0
$ForestRODCCount = 0
$ForestGCCount = 0
$ForestGPOCount = 0
$ForestOUsCount = 0
$ForestDomainDCCountReportTable = @()

ForEach ($ForestDomainItem in $ForestDomains)
 {
    $DomainOUs = @()
    $AllDomainTopLevelOUsItemRecords = @()
    Write-Output "Getting AD Object data from $ForestDomainItem"
    [string]$DomainDC = (Get-ADDomainController -discover -domain $ForestDomainItem).HostName

    # Determine the AD Domain Functional Level
    $ADDomainInfo = Get-ADDomain $ForestDomainItem
    [string]$DomainFunctionalLevel = ($ADDomainInfo.DomainMode) 
    [string]$ForestDomainFunctionalLevels += $DomainFunctionalLevel + ', '

    # Get Domain user statistics
    [array]$DomainUsers = Get-ADUser -filter * -prop AdminCount -Server $DomainDC   
    [int]$ForestUserCount += $DomainUsers.count 

    # Get Domain administrator statistics
    [array]$DomainADAdminArray = Get-ADGroup 'Administrators' -Server $DomainDC | Get-ADGroupMember -Server $DomainDC -Recursive
    [int]$ForestADAdminCount += $DomainADAdminArray.Count

    # Get Domain computer statistics
    [array]$DomainComputers = Get-ADComputer -filter * -Server $DomainDC -Properties OperatingSystem
    [Array]$DomainComputerRunningWindowsServerArray = $DomainComputers | Where { ($_.OperatingSystem -Like "*Windows*") -AND ($_.OperatingSystem -Like "*Server*") }
    [int]$ForestComputerRunningWindowsServerCount += $DomainComputerRunningWindowsServerArray.Count

    [Array]$DomainComputerRunningWindowsWorkstationArray = $DomainComputers | Where { ($_.OperatingSystem -Like "*Windows*") -AND ($_.OperatingSystem -NotLike "*Server*") }
    [int]$ForestComputerRunningWindowsWorkstationCount += $DomainComputerRunningWindowsWorkstationArray.Count

    # Get Domain DC statistics
    [int]$DomainDCCount = 0
    [int]$DomainDCCount = ((Get-ADDomain -Server $DomainDC | select ReplicaDirectoryServers).ReplicaDirectoryServers).count
    [int]$ForestDCCount += $DomainDCCount
    
    # Get Domain RODC statistics
    [int]$DomainRODCCount = 0
    [int]$DomainRODCCount = ((Get-ADDomain -Server $DomainDC | select ReadOnlyReplicaDirectoryServers).ReadOnlyReplicaDirectoryServers).count
    [int]$ForestRODCCount += $DomainRODCCount
    
    # Get Domain GC statistics
    [array]$DomainGCArray = Get-ADDomainController -Filter {IsGlobalCatalog -eq $True} -Server $DomainDC
    [int]$ForestGCCount += $DomainGCArray.Count

    # Get Domain GPO statistics
    [int]$DomainGPOCount = 0
    [int]$DomainGPOCount = (Get-GPO -domain $ForestDomainItem -ALL).Count
    [int]$ForestGPOCount += $DomainGPOCount

    # Get Domain OU statistics
    [int]$DomainOUsCount = 0
    [int]$DomainOUsCount = (Get-ADOrganizationalUnit -Filter * -Server $DomainDC).count  
    [int]$ForestOUsCount += $DomainOUsCount

    $ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest", $ForestDNSName)
    $Forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)

    # Get Domain DC OS statistics
    Write-Output "Discovering DCs in $ForestDomainItem"
    $DomainDCOSCountHashTable = @{}
                
    [array]$DomainDCs = Get-ADDomainController -filter * -Server $DomainDC
    IF (!$DomainDCs)
     { 
        $DomainDCs = @()
        [array]$DomainDCHostNameArray = ((Get-ADDomain -Server $DomainDC | select ReplicaDirectoryServers).ReplicaDirectoryServers) 
        ForEach ($DomainDCHostNameArrayItem in $DomainDCHostNameArray)
         {
            $DomainDCHostNameArrayItemName = ($DomainDCHostNameArrayItem -Split('\.'))[0]
            [array]$DomainDCs += Get-ADDomainController $DomainDCHostNameArrayItemName
         }
     }

    [int]$DomainDCsCount = 0
    [int]$DomainDCsCount = $DomainDCs.Count
    [array]$ForestDCs += $DomainDCs  

    ForEach ($WindowsServerOperatingSystemArrayItem in $WindowsServerOperatingSystemArray)
     {
        [int]$DCOSCount = 0
        $WindowsServerOSDesignatorArray = $WindowsServerOperatingSystemArrayItem -Split(' ')
        $WindowsServerOSDesignator = $WindowsServerOSDesignatorArray[2]
        IF ($WindowsServerOperatingSystemArrayItem -like "*R2*")
         {
            $WindowsServerOSDesignator = $WindowsServerOSDesignator + " R2"
            [array]$DomainDCbyOS = $DomainDCs | where { $_.OperatingSystem -like "*$WindowsServerOSDesignator*" }
            [int]$DCCount = $DomainDCbyOS.count
         }
        ELSE
         {
            [array]$DomainDCbyOS = $DomainDCs | where { ($_.OperatingSystem -match "$WindowsServerOSDesignator") -AND ($_.OperatingSystem -notmatch "R2") }
            [int]$DCCount = $DomainDCbyOS.count
         }

        IF ($DCCount -ge 1)
         {  
            $DomainDCCountReportTable = New-Object PSObject 
            $DomainDCCountReportTable | add-member -membertype NoteProperty -name Domain -Value $ForestDomainItem
            $DomainDCCountReportTable | add-member -membertype NoteProperty -name OperatingSystem -Value $WindowsServerOperatingSystemArrayItem
            $DomainDCCountReportTable | add-member -membertype NoteProperty -name Count -Value $DCCount

            [array]$ForestDomainDCCountReportTable += $DomainDCCountReportTable
         }
        $DomainDCbyOS = @()
     }
 }

[string]$ForestDomainFunctionalLevels = $ForestDomainFunctionalLevels.Substring(0,$ForestDomainFunctionalLevels.Length-2)


$ForestReportText1 = 
@"
$ForestDomainItem Forest Report        
--------------------------------------          

2. Year AD Forest was created: $ForestCreationDate 

3. Forest Total Domain Count:  $ForestDomainCount 

4. Forest Total Organizational Unit (OU) Count: $ForestOUsCount
              
Domain Controllers: $ForestDCCount  
Read-Only Domain Controllers: $ForestRODCCount

5. Total number of Domain Controllers by Operating System in each production forest:      
"@


$ForestReportText2 = 
@"
6. Forest & Domain Functional Levels
    Forest Functional Level: $ForestFunctionalLevel 
    Domain Functional Levels: $ForestDomainFunctionalLevels

8. Total Number of Forest Users: $ForestUserCount   
9. Total Number of Forest Workstations: $ForestComputerRunningWindowsWorkstationCount 
10. Total Number of Forest Servers: $ForestComputerRunningWindowsServerCount 
11. Total Number of Forest Active Directory Admins: $ForestADAdminCount
12. Total Number of Forest AD Sites: $TotalSitesInForest 
13. Total Number of Group Policy Objects (GPOs): $ForestGPOCount  

                                                
"@

Write-Output ""
Write-Output ""

$ForestReportText1 
$ForestDomainDCCountReportTable | select * | format-table -AutoSize
$ForestReportText2

Write-Output "Script Completed."
Stop-Transcript
