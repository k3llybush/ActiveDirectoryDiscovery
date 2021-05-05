$Counts = @()

$final_local = "$env:userprofile\Desktop\ADData\Counts\";

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

$Domain = (Get-ADDomain).DistinguishedName
$Users = (get-aduser -Filter *).count
$UsersEnabled = (get-aduser -filter *|Where-Object {$_.enabled -eq "True"}).count 
$UsersDisabled = (get-aduser -filter *|Where-Object {$_.enabled -ne "False"}).count
$Groups = (Get-ADGroup -Filter *).count
$Contacts = (Get-ADObject -Filter 'ObjectClass -eq "contact"' -Searchbase (Get-ADDomain).distinguishedName).count
$Computers = (Get-ADComputer -Filter *).count
$Workstations = (Get-ADComputer -LDAPFilter "(&(objectClass=Computer)(!operatingSystem=*server*)(operatingSystem=*windows*))" -Searchbase (Get-ADDomain).distinguishedName).count
$Servers = (Get-ADComputer -LDAPFilter "(&(objectClass=Computer)(operatingSystem=*server*))" -Searchbase (Get-ADDomain).distinguishedName).count
$OtherComputer = (Get-ADComputer -LDAPFilter "(&(objectClass=Computer)(!operatingSystem=*server*)(!operatingSystem=*windows*))" -Searchbase (Get-ADDomain).distinguishedName).count
$GPOs = (Get-GPO -All).count
$OUs = (Get-ADOrganizationalUnit -Filter *).count

Write-Host "Users =          "$Users
Write-Host "Users Enabled =  "$UsersEnabled
Write-Host "Users Disabled = "$UsersDisabled
Write-Host "Groups =         "$Groups
Write-Host "Contacts =       "$Contacts
Write-Host "Computers =      "$Computers
Write-Host "Workstions =     "$Workstations
Write-Host "Servers =        "$Servers
Write-Host "Other Computers ="$OtherComputer
Write-Host "Domain =         "$Domain
Write-Host "GPOs =           "$GPOs
Write-Host "OUs =            "$OUs

#Use New-Object and Add-Member to create an object
$Counts = New-Object System.Object
$Counts | Add-Member -MemberType NoteProperty -Name "Date" -Value $Date
$Counts | Add-Member -MemberType NoteProperty -Name "Users" -Value $Users
$Counts | Add-Member -MemberType NoteProperty -Name "Users Enabled" -Value $UsersEnabled
$Counts | Add-Member -MemberType NoteProperty -Name "Users Disabled" -Value $UsersDisabled
$Counts | Add-Member -MemberType NoteProperty -Name "Groups" -Value $Groups
$Counts | Add-Member -MemberType NoteProperty -Name "Contacts" -Value $Contacts
$Counts | Add-Member -MemberType NoteProperty -Name "Computers" -Value $Computers
$Counts | Add-Member -MemberType NoteProperty -Name "Workstations" -Value $Workstations
$Counts | Add-Member -MemberType NoteProperty -Name "Servers" -Value $Servers
$Counts | Add-Member -MemberType NoteProperty -Name "Other Computer" -Value $OtherComputer
$Counts | Add-Member -MemberType NoteProperty -Name "GPOs" -Value $GPOs
$Counts | Add-Member -MemberType NoteProperty -Name "OUs" -Value $OUs
$Counts | Add-Member -MemberType NoteProperty -Name "Domain" -Value $Domain

#Finally, use Export-Csv to export the data to a csv file
$Counts | Export-Csv -NoTypeInformation -Path "$final_local\Counts.csv" -Append


