$local = Get-Location;
$final_local = "$env:userprofile\Desktop\ADData\allADUsers";

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

Search-ADAccount -PasswordNeverExpires | Export-Csv "$final_local\neverexpires.csv" -NoTypeInformation