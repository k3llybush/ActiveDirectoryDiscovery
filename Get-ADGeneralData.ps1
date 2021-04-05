$Date = Get-Date -Format d
$Counts = @()

#Clear-Variable "final_local"

$local = Get-Location;
$final_local = "$env:userprofile\Desktop\ADData\DomainControllers\";

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


Get-ADForest > $final_local\Forest.txt
Get-ADDomain > $final_local\Domain.txt
Get-ADDefaultDomainPasswordPolicy > $final_local\DD-PWD-Policy.txt
Get-ADFineGrainedPasswordPolicy -Filter * >$final_local\Fine-Grained.txt
Get-ADOptionalFeature -Filter * > $final_local\ADOptionalFeature.txt
Get-ADTrust -Filter * > $final_local\ADTrust.txt

