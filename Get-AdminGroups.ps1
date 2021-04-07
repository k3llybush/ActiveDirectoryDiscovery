$final_local = "$env:userprofile\Desktop\ADData\AdminGroups\$date";

$date = get-date -format M.d.yyyy
$local = Get-Location;

if (!$local.Equals("C:\")) {
    Set-Location "C:\";
    if ((Test-Path $final_local) -eq 0) {
        mkdir $final_local;
        Set-Location $final_local;
    }

    ## if path already exists
    ## DB Connect
    elseif ((Test-Path $final_local) -eq 1) {
        Set-Location $final_local;
        Write-Output $final_local;
    }
}

$DomainSID = (Get-ADDomain).DomainSID

$Groups = get-adgroup -Filter *
foreach ($g in $Groups) {
    $g = $g.sid.Value 
    $3 = $g.substring($g.length - 4).trimstart("-")

    if ([int]$3 -gt "2000") {
    }
    else {
        $sid = $DomainSID.Value + "-" + $3
        $Group = Get-ADGroup -Filter { SID -eq $sid }
        $Name = $Group.Name
        Get-ADGroupMember $Group  | Select-Object name, distinguishedName | Out-File "$final_local\$name.txt"
    }

}