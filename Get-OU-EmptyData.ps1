$local = Get-Location;
$final_local = "$env:userprofile\Desktop\ADData\OUData\";

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


Get-ADOrganizationalUnit -Filter * -Properties * | 
Where-Object {-not ( Get-ADObject -Properties * -Filter * -SearchBase $_.Distinguishedname `
    -SearchScope OneLevel -ResultSetSize 1 )} | select name,distinguishedName,Created | `
    Select-Object | export-csv "$final_local\empty-OUs.txt" -NoTypeInformation



