$local = Get-Location;
$final_local = "$env:userprofile\Desktop\ADData\AllADUsers";

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


Get-ADUser -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} –Properties “DisplayName”, “msDS-UserPasswordExpiryTimeComputed” |

Select-Object -Property “Displayname”,@{Name=“ExpiryDate”;Expression={[datetime]::FromFileTime($_.“msDS-UserPasswordExpiryTimeComputed”)}} | Export-Csv "$final_local\password-expires-date.csv" -NoTypeInformation
