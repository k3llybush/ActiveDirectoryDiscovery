$final_local = "$env:userprofile\Desktop\ADData\AdminGroups\$date";

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


Get-ADGroupMember "Account Operators" | select name,distinguishedName | Out-File "$final_local\AccountOperators.txt"
Get-ADGroupMember "Administrators" | select name,distinguishedName | Out-File "$final_local\builtinAdministrators.txt"
Get-ADGroupMember "Backup Operators" | select name,distinguishedName | Out-File "$final_local\BackupOperators.txt"
Get-ADGroupMember "Print Operators" | select name,distinguishedName | Out-File "$final_local\PrintOperators.txt"
Get-ADGroupMember "Remote Desktop Users" | select name,distinguishedName | Out-File "$final_local\RDPUsers.txt"
Get-ADGroupMember "Server Operators" | select name,distinguishedName | Out-File "$final_localServerOperators.txt"
Get-ADGroupMember "DNSAdmins" | select name,distinguishedName | Out-File "$final_local\DNSAdmins.txt"
Get-ADGroupMember "Domain ADmins" | select name,distinguishedName | Out-File "$final_local\DA.txt"

