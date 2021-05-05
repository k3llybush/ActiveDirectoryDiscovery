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

Get-ADForest > $final_local\Forest.txt
Get-ADDomain > $final_local\Domain.txt

Get-ADDomainController -Filter * | Export-Csv "$final_local\Get-ADDomainController.csv" -NoTypeInformation
Get-ADDefaultDomainPasswordPolicy | Export-Csv "$final_local\Get-ADDefaultDomainPasswordPolicy.csv" -NoTypeInformation
Get-ADFineGrainedPasswordPolicy -Filter * | Export-Csv "$final_local\Get-ADFineGrainedPasswordPolicy.csv" -NoTypeInformation
Get-ADOptionalFeature -Filter * | Export-Csv "$final_local\Get-ADOptionalFeature.csv" -NoTypeInformation 
Get-ADTrust -Filter * -Properties * | Export-Csv "$final_local\Get-ADTrust.csv" -NoTypeInformation 

Get-ADAuthenticationPolicy -LDAPFilter '(name=AuthenticationPolicy*)' | Export-Csv "$final_local\Get-ADAuthenticationPolicy.csv" -NoTypeInformation
Get-ADAuthenticationPolicySilo -Filter '(Name -like "*AuthenticationPolicySilo*")' | Export-Csv "$final_local\Get-ADAuthenticationPolicySilo.csv" -NoTypeInformation
Get-ADCentralAccessPolicy -Filter * | Export-Csv "$final_local\Get-ADCentralAccessPolicy.csv" -NoTypeInformation
Get-ADCentralAccessRule -Filter * | Export-Csv "$final_local\Get-ADCentralAccessRule.csv" -NoTypeInformation 
Get-ADClaimTransformPolicy -Filter * | Export-Csv "$final_local\Get-ADClaimTransformPolicy .csv" -NoTypeInformation
Get-ADClaimType -Filter * | Export-Csv "$final_local\Get-ADClaimType.csv" -NoTypeInformation

Get-ADOrganizationalUnit -Filter * -Properties * | Export-Csv "$final_local\Get-ADOrganizationalUnit.csv" -NoTypeInformation
Get-ADReplicationSite -Filter * | Export-Csv "$final_local\Get-ADReplicationSite.csv" -NoTypeInformation
Get-ADReplicationSubnet -Filter * | Export-Csv "$final_local\Get-ADReplicationSubnet.csv" -NoTypeInformation
Get-ADReplicationSiteLink -Filter * | Export-Csv "$final_local\Get-ADReplicationSiteLink.csv" -NoTypeInformation

$DNSRoot = Get-ADDomain

Resolve-DnsName -Name "_ldap._tcp.$($DNSRoot.DNSRoot)" -Type srv | Sort-Object nametarget, name, type | Export-Csv "$final_local\_ldap_tcp.csv" -NoTypeInformation
Resolve-DnsName -Name "_kerberos._tcp.$($DNSRoot.DNSRoot)" -Type srv | Sort-Object nametarget, name, type | Export-Csv "$final_local\_kerberos_tcp.csv" -NoTypeInformation