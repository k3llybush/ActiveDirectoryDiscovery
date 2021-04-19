$final_local = "$env:userprofile\Desktop\ADData\AdminGroups\$date";

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

$DomainSID = (Get-ADDomain).DomainSID
$i = 0

foreach ($_ in 1..2001) {
    $i++
    
    try {
        $sid = $DomainSID.Value + "-$i"
        $Group = Get-ADGroup -Filter { SID -eq $sid }
        $Name = $Group.Name
        Get-ADGroupMember $Name  | Select-Object name, distinguishedName | Out-File "$final_local\$name.txt"
    }
    catch {}
}