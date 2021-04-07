$final_local = "$env:userprofile\Desktop\ADData\DomainControllers\";

#$date = get-date -format M.d.yyyy
$local = Get-Location;

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

Get-ADForest > $final_local\Forest.txt
Get-ADDomain > $final_local\Domain.txt
Get-ADDefaultDomainPasswordPolicy > $final_local\DD-PWD-Policy.txt
Get-ADFineGrainedPasswordPolicy -Filter * >$final_local\Fine-Grained.txt
Get-ADOptionalFeature -Filter * > $final_local\ADOptionalFeature.txt
Get-ADTrust -Filter * > $final_local\ADTrust.txt

