﻿<#
  This script will enumerate all enabled user accounts in a Domain, calculate their estimated Token
  Size and create two reports in CSV format:
  1) A report of all users with an estimated token size greater than or equal to the number defined
     by the $TokensSizeThreshold variable.
  2) A report of the top x users as defined by the $TopUsers variable.

  Syntax:

  - To run the script against all enabled user accounts in the current domain:
      Get-TokenSizeReport.ps1

  - To run the script against all enabled user accounts of a trusted domain:
      Get-TokenSizeReport.ps1 -TrustedDomain:mytrusteddomain.com

  - To run the script against 1 user account:
      Get-TokenSizeReport.ps1 -AccountName:<samaccountname>

  - To run the script against 1 user account of a trusted domain:
      Get-TokenSizeReport.ps1 -AccountName:<samaccountname> -TrustedDomain:mytrusteddomain.com

  Script Name: Get-TokenSizeReport.ps1
  Release 2.7
  Written by Jeremy Saunders (jeremy@jhouseconsulting.com) 13/12/2013
  Modified by Jeremy Saunders (jeremy@jhouseconsulting.com) 19/12/2016

  Original script was derived from the CheckMaxTokenSize.ps1 written by Tim Springston [MS] on 7/19/2013
  http://gallery.technet.microsoft.com/scriptcenter/Check-for-MaxTokenSize-520e51e5

  Re-wrote the script to be more efficient and provide a report for all users in the
  Domain.

  References:
  - Microsoft KB327825: Problems with Kerberos authentication when a user belongs to many groups
    http://support.microsoft.com/kb/327825
  - Microsoft KB243330: Well-known security identifiers in Windows operating systems
    http://support.microsoft.com/kb/243330
  - Microsoft KB328889: Users who are members of more than 1,015 groups may fail logon authentication
    http://support.microsoft.com/kb/328889
  - Microsoft KB938118: How to use Group Policy to add the MaxTokenSize registry entry to multiple computers
    http://support.microsoft.com/kb/938118
  - Microsoft Blog: Managing Token Bloat:
    http://blogs.dirteam.com/blogs/sanderberkouwer/archive/2013/05/22/common-challenges-when-managing-active-directory-domain-services-part-2-unnecessary-complexity-and-token-bloat.aspx

  The calculation in this script is an estimation based on the formula documented under KB327825. I prefer
  to review any account with a token size over 6K, hence why I default the $TokensSizeThreshold variable
  to 6000.

  A user's access token is increased in 4K blocks. The default size of a user's access token is 4K. Once
  a user goes over this amount Windows does not increment the size of the token by the amount needed for
  each additional SID added. Instead Windows allocates another 4K of memory, thus doubling the size of
  the access token to 8K. And again once the size of the access token breaches 8K it will again increase
  by a further 4K to 12K. And so on.

  So given that the token size calculation is an estimation, and the access token is increased in blocks
  of 4K, it is quite conceivable that an estimated calculated token size of between 6000 and 12000 can
  actually breach the 12K token size limit and cause problems in some environments. There is a real lack
  of understanding here.

  Although it's not documented in KB327825 or any other Microsoft references, I also add the number of
  global groups outside the user's account domain that the user is a member of to the "d" calculation of
  the TokenSize. Whilst the Microsoft methodology is to add universal groups from other domains, it is
  possible to add global groups too. Therefore it's important to capture this and correctly include it in
  the calculation.

  My calculations consider the BuiltIn groups as Domain Local groups, which means I'm allowing 40 bytes
  per BuiltIn group that the user is a member of. Others seem to only allow 8 bytes in their calculations.
  However depending on the length of the SID, BuiltIn groups are actually either 16 or 28 bytes in reality.
  Therefore, whilst I may be overcompensating for some of the groups, others are always underestimating.

  If we wanted to be as accurate as possible we can calculate the byte length of each SID and then add a
  further 16 bytes for associated attributes and information. Most user and group SIDs are 28 bytes in
  length.

  Refer to the following thread to get a full understanding how this needs to be calculated for complete
  acuracy. You also need to understand how the token is managed in 4KB blocks. It starts to make sense
  when you tie together the comments from Paul Bergson, Richard Mueller and Marcin Polich:
  - https://social.technet.microsoft.com/Forums/windowsserver/en-US/a7035740-355f-4a8c-a434-27583e3b6075/probably-a-maxtokensize-problem-but?forum=winserverDS

  There is some further good information on SIDs published by Philipp Fockeler:
  - http://www.selfadsi.org/deep-inside/microsoft-sid-attributes.htm

  For users with large tokens consider reducing direct and transitive (nested) group memberships.
  Larger environments that have evolved over time also have a tendancy to suffer from Circular Group
  Nesting and both user and group SIDHistory.

  The SIDHistory of the user and the group accounts are included in the calculation of the token size,
  which is why it's important to clean up SIDHistory after a domain migration. From experience I find
  that this rarely happens.

  It's important to note that with Windows 2012 Active Directory, compression enhancements were added to
  the KDC functionality. Therefore the formula used in this script does not apply to when the domain
  functionality level is Windows 2012 and above.

  On the odd ocasion I was receiving the following error:
  - Exception calling "FindByIdentity" with "2" argument(s): "Multiple principals contain
    a matching Identity."
  - There seemed to be a known bug in .NET 4.x when passing two arguments to the FindByIdentity() method.
  - The solution was to either use a machine with .NET 3.5 or re-write the script to pass
    three arguments as per the Get-UserPrincipal function provided in the following Scripting Guy article:
    - http://blogs.technet.com/b/heyscriptingguy/archive/2009/10/08/hey-scripting-guy-october-8-2009.aspx
    This function passes the Context Type, FQDN Domain Name and Parent OU/Container.
  - Other references:
    - http://richardspowershellblog.wordpress.com/2008/05/27/account-management-member-of/
    - http://www.powergui.org/thread.jspa?threadID=20194

  I have also seen the following error:
  - Exception calling "GetAuthorizationGroups" with "0" argument(s): "An error (1301) occurred while
    enumerating the groups. The group's SID could not be resolved."
  - Other references:
    - http://richardspowershellblog.wordpress.com/2008/05/27/account-management-member-of/
    - https://groups.google.com/forum/#!topic/microsoft.public.adsi.general/jX3wGd0JPOo
    - http://lucidcode.com/2013/02/18/foreign-security-groups-in-active-directory/

  Added the tokenGroups attribute to get all nested groups as I could not achieve 100% reliability using
  the GetAuthorizationGroups() method. Could not afford for it to start failing after running for hours
  in large environments.
  - References:
    - http://www.msxfaq.de/code/tokengroup.htm
    - http://www.msxfaq.de/tools/dumpticketsize.htm

  There are important differences between using the GetAuthorizationGroups() method versus the tokenGroups
  attribute that need to be understood. Aside from the unreliability of GetAuthorizationGroups(), when push
  comes to shove you get different results depending on which method you use, and what you want to achieve.
    - The tokenGroups attribute only contains the actual "Active Directory" principals, which are groups and
      siDHistory (from both the user and groups that th user is a member of).
    - However, whilst tokenGroups contains transitive groups (groups within the same forest), it does not
      reveal cross-forest/domain group memberships. The tokenGroups attribute is constructed by Active
      Directory on request, and this depends on the availability of a Global Catalog server:
      http://msdn.microsoft.com/en-us/library/ms680275(VS.85).aspx
    - The GetAuthorizationGroups() method also returns the well-known security identifiers of the local
      system (LSALogonUser) for the user running the script, which will include groups such as:
      - Everyone (S-1-1-0)
      - Authenticated Users (S-1-5-11)
      - This Organization (S-1-5-15)
      - Low Mandatory Level (S-1-16-4096)
      This will vary depending on where you're running the script from and in what user context. The result
      is still consistent, as it adds the same overhead to each user. But this is misleading.
    - GetAuthorizationGroups() will return cross-forest/domain group memberships, but cannot resolve them
      because they contain a ForeignSecurityPrincipal. It therefore fails as documented above.
    - GetAuthorizationGroups() does not contain siDHistory.

  In my view you would always use the tokenGroups attribute to collate a consistent and accurate user report
  across the environment, whereas the GetAuthorizationGroups() method could be used in a logon script to
  calucate the token of the user together with the system they are logging on to. The actual calculation of
  the token size adds the estimated value for ticket overhead anyway, hence the reason why using the tokenGroups
  attribute provides a consistent result for all users.
  If you wanted an accurate token size per user per system and GetAuthorizationGroups() method continues to
  prove to be unreliable, you could use the tokenGroups attribute together with the addition of the output
  from the "whoami /groups" command to get all the well-known groups and label needed to calculate the
  complete local token.

  Microsoft also has a tool called Tokensz.exe that could also be used in a logon script. It can be downloded
  from here: http://www.microsoft.com/download/en/details.aspx?id=1448

  To be completed:
  - Implement a garbage collection process to reduce memory usage:
    - http://powershell.com/cs/blogs/tips/archive/2015/05/15/get-memory-consumption.aspx
    - https://dmitrysotnikov.wordpress.com/2012/02/24/freeing-up-memory-in-powershell-using-garbage-collector/
  - Some further research and testing needs to be completed with the code that retrieves the tokenGroups
    attribute to validate performance between the GetInfoEx method or RefreshCache method.
  - Investigate if using the WindowsIdentity.Groups Property to get the users token is a better approach.

#>

#-------------------------------------------------------------
param(
      [String]$AccountName,
      [String]$TrustedDomain,
      [String]$InputFile
     )

# Set Powershell Compatibility Mode
Set-StrictMode -Version 2.0

# Enable verbose, warning and error mode
$VerbosePreference = 'Continue'
$WarningPreference = 'Continue'
$ErrorPreference = 'Continue'

#-------------------------------------------------------------

# Set this to the OU structure where the you want to search to
# start from. Do not add the Domain DN. If you leave it blank,
# the script will start from the root of the domain.
$OUStructureToProcess = ""

# Set the search scope:
# - SUBTREE is the defined OU and all child OUs (including their
#   children, etc)
# - ONELEVEL is the defined container
$SearchScope = "SUBTREE"

# Set the name of the OU(s) you want to exclude. Use the full OU
# structure minus the Domain DN.
$ExcludeOUs = @("OU=ExcludeMe","OU=ExcludeMeToo,OU=People")

# Set this value to the number of users with large tokens that
# you want to report on.
$TopUsers = 200

# Set this to the size in bytes that you want to capture the user
# information for the report.
$TokensSizeThreshold = 1200

# Set this value to true if you want to output to the console
$ConsoleOutput = $True

# Set this value to true if you want a summary output to the
# console when the script has completed.
$OutputSummary = $True

# Set this to the delimiter for the CSV output
$Delimiter = ","

# Set this value to true if you want to see the progress bar.
$ProgressBar = $True

#-------------------------------------------------------------

write-verbose "This script is running under PowerShell version $($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)"

$invalidChars = [io.path]::GetInvalidFileNamechars() 
$datestampforfilename = ((Get-Date -format s).ToString() -replace "[$invalidChars]","-")

# Get the script path
$ScriptPath = {Split-Path -parent "$env:userprofile\Desktop\ADData\AllADUsers\*.*"}
$ReferenceFile = $(&$ScriptPath) + ".\KerberosTokenSizeReport-$($datestampforfilename).csv"
$ReferenceFileTopUsers = $(&$ScriptPath) + ".\KerberosTokenSizeReport-TopUsers-$($datestampforfilename).csv"

if (Test-Path -path $ReferenceFile) {
  remove-item $ReferenceFile -force -confirm:$false
}

if (Test-Path -path $ReferenceFileTopUsers) {
  remove-item $ReferenceFileTopUsers -force -confirm:$false
}

if ([String]::IsNullOrEmpty($TrustedDomain)) {
  # Get the Current Domain Information
  $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
} else {
  $context = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext("domain",$TrustedDomain)
  Try {
    $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($context)
  }
  Catch [exception] {
    $Host.UI.WriteErrorLine("ERROR: $($_.Exception.Message)")
    Exit
  }
}

# Get AD Distinguished Name
$DomainDistinguishedName = $Domain.GetDirectoryEntry() | select -ExpandProperty DistinguishedName  

If ($OUStructureToProcess -eq "") {
  $ADSearchBase = $DomainDistinguishedName
} else {
  $ADSearchBase = $OUStructureToProcess + "," + $DomainDistinguishedName
}

$garbageCounter = 0
$arrayoftopusers = @()
$TotalUsersProcessed = 0
$UserCount = 0
$GroupCount = 0
$LargestTokenSize = 0
$TotalTokensLessThan8K = 0
$TotalTokensBetween8and12K = 0
$TotalTokensGreaterThan12K = 0
$TotalTokensGreaterThan48K = 0
$TotalExcludedUsers = 0

$UseInputFile = $False
If (-not [String]::IsNullOrEmpty($InputFile)) {
  $InputFile = $(&$ScriptPath) + "\$InputFile"
  If (Test-Path $InputFile) {
    $UseInputFile = $True
  }
}

If ([String]::IsNullOrEmpty($AccountName)) {
  # Create an LDAP search for all enabled users
  $ADFilter = "(&(objectClass=user)(objectcategory=person)(!userAccountControl:1.2.840.113556.1.4.803:=2))"
  $ProcessSingleAccount = $False
} Else {
  # Create an LDAP search for a simple user
  $ADFilter = "(&(objectClass=user)(objectcategory=person)(samaccountname=$AccountName))"
  $ProcessSingleAccount = $True
  $TokensSizeThreshold = 65335
  $OutputSummary = $False
}

# There is a known bug in PowerShell requiring the DirectorySearcher
# properties to be in lower case for reliability.
$ADPropertyList = @("distinguishedname","samaccountname","useraccountcontrol","objectsid","sidhistory","primarygroupid","lastlogontimestamp","memberof")
$ADScope = $SearchScope
$ADPageSize = 1000
$ADSearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($ADSearchBase)") 
$ADSearcher = New-Object System.DirectoryServices.DirectorySearcher 
$ADSearcher.SearchRoot = $ADSearchRoot
$ADSearcher.PageSize = $ADPageSize 
$ADSearcher.Filter = $ADFilter 
$ADSearcher.SearchScope = $ADScope
if ($ADPropertyList) {
  foreach ($ADProperty in $ADPropertyList) {
    [Void]$ADSearcher.PropertiesToLoad.Add($ADProperty)
  }
}
Try {
  write-host " "
  If ([String]::IsNullOrEmpty($AccountName)) {
    write-verbose "Please be patient whilst the script retrieves all enabled user objects and specified attributes..."
  } Else {
    write-verbose "Retrieving the user object and specified attributes..."
  }
  write-host " "
  $UserCount = $ADSearcher.FindAll().Count
}
Catch {
  $UserCount = 0
  $Host.UI.WriteErrorLine("ERROR: The $ADSearchBase structure cannot be found!")
}
Finally {
  # Dispose of the search and results properly to avoid a memory leak
  $ADSearcher.Dispose() | Out-Null
  [System.GC]::Collect() | Out-Null
}

if ($UserCount -ne 0) {
  # The ForEach-Object cmdlet processes each item in turn as it is passed through the pipeline
  # whereas foreach generates the whole collection first. So this should alleviate memory issues.
  $ADSearcher.Findall() | ForEach-Object {
    #$_.Properties
    #$_.Properties.propertynames
    $lastLogonTimeStamp = ""
    $lastLogon = ""
    $UserDN = $_.Properties.distinguishedname[0]
    $samAccountName = $_.Properties.samaccountname[0]
    $ParentOU = $UserDN -split '(?<![\\]),'
    $ParentOU = $ParentOU[1..$($ParentOU.Count-1)] -join ','

    $TotalUsersProcessed ++
    If ($ProgressBar) {
      Write-Progress -Activity "Processing $($UserCount) Users" -Status ("Count: $($TotalUsersProcessed) - Username: {0}" -f $samAccountName) -PercentComplete (($TotalUsersProcessed/$UserCount)*100)
    }

    $ExcludeOUMatch = $False
    If ($ExcludeOUs -eq "" -OR ($ExcludeOUs | Measure-Object).Count -eq 0) {
      $ExcludeOUMatch = $False
    } Else {
      ForEach ($ExcludeOU in $ExcludeOUs) {
        If ($ParentOU -Like "*$ExcludeOU*") {
          $ExcludeOUMatch = $True
        }
      }
    }

    If ($ExcludeOUMatch -eq $False) {

      If ($(Try{($_.Properties.lastlogontimestamp | Measure-Object).Count -gt 0}Catch{$False})) {
        $lastLogonTimeStamp = $_.Properties.lastlogontimestamp[0]
        $lastLogon = [System.DateTime]::FromFileTime($lastLogonTimeStamp)
        if ($lastLogon -match "1/01/1601") {$lastLogon = "Never logged on before"}
      } else {
        $lastLogon = "Never logged on before"
      }
      $OU = $_.GetDirectoryEntry().Parent
      $OU = $OU -replace ("LDAP:\/\/","")

      # Get user SID
      $arruserSID = New-Object System.Security.Principal.SecurityIdentifier($_.Properties.objectsid[0], 0)
      $userSID = $arruserSID.Value

      # Get the SID of the Domain the account is in
      $AccountDomainSid = $arruserSID.AccountDomainSid.Value

      # Get User Account Control & Primary Group by binding to the user account
      # ADSI Requires that / Characters be Escaped with the \ Escape Character
      $UserDN = $UserDN.Replace("/", "\/")
      $objUser = [ADSI]("LDAP://" + $UserDN)
      If ($(Try{($objUser.useraccountcontrol | Measure-Object).Count -gt 0}Catch{$False})) {
        $UACValue = $objUser.useraccountcontrol[0]
      } else {
        $UACValue = ""
      }

      $Enabled = $True
      if (($UACValue -bor 0x0002) -eq $UACValue) {
        $Enabled = $False
      }

      $TrustedforDelegation = $false
      if ((($UACValue -bor 0x80000) -eq $UACValue) -OR (($UACValue -bor 0x1000000) -eq $UACValue)) {
        $TrustedforDelegation = $true
      }

      $primarygroupID = $objUser.primarygroupid
      If ($(Try{$primarygroupID -ne $NULL}Catch{$False})) {
        # Primary group can be calculated by merging the account domain SID and primary group ID
        $primarygroupSID = $AccountDomainSid + "-" + $primarygroupID.ToString()
        $primarygroup = [adsi]("LDAP://<SID=$primarygroupSID>")
        $primarygroupname = $primarygroup.name[0]
      } else {
        $primarygroupname = "NULL"
      }
      $objUser = $null

      # Get User SID history
      $UserSIDHistoryCount = 0
      if ($(Try{$_.Properties.sidhistory -ne $null}Catch{$False})) {
        foreach ($sidhistory in $_.Properties.sidhistory) {
          $SIDHistObj = New-Object System.Security.Principal.SecurityIdentifier($sidhistory, 0)
          #write-verbose "$($SIDHistObj.Value) is in the SIDHistory."
          $UserSIDHistoryCount++
        }
        $SIDHistObj = $null
      }

      $UserAccount = [ADSI]"$($_.Path)"
      $tokenGroupsMethod1 = $true
      $tokenGroupsMethod2 = $false
      If ($tokenGroupsMethod1) {
        $UserAccount.GetInfoEx(@("tokenGroups"),0) | Out-Null
        $ErrorActionPreference = "continue"
        $error.Clear()
        $tokengroups = $UserAccount.GetEx("tokengroups")
        if ($Error) {
          Write-Warning "  Tokengroups not readable"
          $tokengroups=@()   #empty enumeration
        }
      }
      If ($tokenGroupsMethod2) {
        # Rebuild the tokenGroups attribute for the user, which is a dynamic attribute not part
        # of the schema.
        $UserAccount.psbase.refreshCache(@("TokenGroups"))
        $tokengroups = $UserAccount.psbase.Properties.Item("tokenGroups")
      }
      $UserPrincipalOutput = @()
      $PrincipalCount = 0
      $GroupCount = 0
      $SecurityDomainLocalScope = 0
      $SecurityGlobalInternalScope = 0
      $SecurityGlobalExternalScope = 0
      $SecurityUniversalInternalScope = 0
      $SecurityUniversalExternalScope = 0
      $ExternalGroupsFound = $false
      $TotalGroupSIDHistoryCount = 0

      foreach($sidByte in $tokengroups) {
        $PrincipalCount++
        $GroupSIDBytes = 0
        $principal = New-Object System.Security.Principal.SecurityIdentifier($sidByte,0)
        $objUserToken = New-Object -TypeName PSObject
        Try{
          $PrincipalAccountName = $principal.Translate([System.Security.Principal.NTAccount])
          $objUserToken | Add-Member -MemberType NoteProperty -Name "Account" -value $PrincipalAccountName.value
        }
        Catch {
          $objUserToken | Add-Member -MemberType NoteProperty -Name "Account" -value "Cannot translate SID to an account"
        }
        $objUserToken | Add-Member -MemberType NoteProperty -Name "SID" -value $principal.Value
        $objUserToken | Add-Member -MemberType NoteProperty -Name "BinaryLength" -value $principal.BinaryLength
        $objUserToken | Add-Member -MemberType NoteProperty -Name "AccountDomainSid" -value $principal.AccountDomainSid

        $grp = [ADSI]("LDAP://<SID=$($principal.Value)>")

        if ($grp.Path -ne $null) {
          $GroupCount++
          $grpdn = $grp.distinguishedName.tostring().ToLower()
          $grouptype = $grp.groupType.value
          $objUserToken | Add-Member -MemberType NoteProperty -Name "GroupType" -value $grouptype

          switch -exact ($grouptype) {
            "-2147483646"   {
                              # Global security scope
                              $groupscope = "Global"
                              if ($principal.AccountDomainSid -eq $AccountDomainSid) 
                              {
                                $SecurityGlobalInternalScope++
                              } else { 
                                # Global groups from others.
                                $SecurityGlobalExternalScope++
                                $ExternalGroupsFound = $true
                              }
                              break
                            } 
            "-2147483644"   { 
                              # Domain Local scope 
                              $groupscope = "DomainLocal"
                              $SecurityDomainLocalScope++
                              break
                            } 
            "-2147483643"   { 
                              # Domain Local BuiltIn scope
                              $groupscope = "Builtin"
                              $SecurityDomainLocalScope++
                              break
                            }
            "-2147483640"   { 
                              # Universal security scope 
                              $groupscope = "Universal"
                              if ($principal.AccountDomainSid -eq $AccountDomainSid)
                              { 
                                $SecurityUniversalInternalScope++ 
                              } else { 
                                # Universal groups from others.
                                $SecurityUniversalExternalScope++ 
                                $ExternalGroupsFound = $true
                              } 
                              break
                            }
            Default         {
                                $groupscope = "Unknown"
                                write-warning "No valid group type found!"
                            }
          }  
          $objUserToken | Add-Member -MemberType NoteProperty -Name "GroupScope" -value $groupscope
  
          $GroupSIDHistoryCount = 0
          If ($(Try{($grp.sidhistory | Measure-Object).Count -gt 0}Catch{$False})) {
            foreach ($groupsidhistory in $grp.sidhistory) {
              (new-object System.Security.Principal.SecurityIdentifier $groupsidhistory,0).Value | out-null
              $GroupSIDHistoryCount++
            }
          }
          $objUserToken | Add-Member -MemberType NoteProperty -Name "SIDHistoryCount" -value $GroupSIDHistoryCount

        } Else {
          # The SID Histoty of a group does not have a group type. Therefore, if the group path equals $null,
          # the principal is a user or group SID History
          $GroupSIDHistoryCount = 0
          $objUserToken | Add-Member -MemberType NoteProperty -Name "GroupType" -value "SIDHistory"
          $objUserToken | Add-Member -MemberType NoteProperty -Name "GroupScope" -value "SIDHistory"
          $objUserToken | Add-Member -MemberType NoteProperty -Name "SIDHistoryCount" -value "N/A"
        }
        $UserPrincipalOutput += $objUserToken
        $objUserToken = $null
        $TotalGroupSIDHistoryCount = $TotalGroupSIDHistoryCount + $GroupSIDHistoryCount
      }
      If ($ProcessSingleAccount) {
        $SingleAccountOutputFile = "$SamAccountName-$domain.csv"
        if (Test-Path -path $SingleAccountOutputFile) {
          remove-item $SingleAccountOutputFile -force -confirm:$false
        }
        if ($PSVersionTable.PSVersion.Major -gt 2) {
          $UserPrincipalOutput | Export-Csv -Path "$SingleAccountOutputFile" -Append -Delimiter $Delimiter -NoTypeInformation -Encoding ASCII
        } Else {
          if (!(Test-Path -path $ReferenceFileTopUsers)) {
            $UserPrincipalOutput | ConvertTo-Csv -NoTypeInformation -Delimiter $Delimiter | Select-Object -First 1 | Out-File -Encoding ascii -filepath "$SingleAccountOutputFile"
          }
          $UserPrincipalOutput | ConvertTo-Csv -NoTypeInformation -Delimiter $Delimiter | Select-Object -Skip 1 | Out-File -Encoding ascii -filepath "$SingleAccountOutputFile" -append -noclobber
        }
      }

      If ($ConsoleOutput) {
        write-verbose "Checking the token of user $SamAccountName in domain $domain"
        # The principals in TokenGroups includes the user SID History and user group SID History
        write-verbose "- There are $PrincipalCount principals in the token. This is made up of:"
        write-verbose "  - $GroupCount groups"
        write-verbose "    - $SecurityDomainLocalScope are domain local security groups"
        write-verbose "    - $SecurityGlobalInternalScope are domain global scope security groups inside the users domain"
        write-verbose "    - $SecurityGlobalExternalScope are domain global scope security groups outside the users domain"
        write-verbose "    - $SecurityUniversalInternalScope are universal security groups inside the users domain"
        write-verbose "    - $SecurityUniversalExternalScope are universal security groups outside the users domain"
        write-verbose "  - $($UserSIDHistoryCount + $TotalGroupSIDHistoryCount) SIDs in the SIDHistory"
        write-verbose "    - $UserSIDHistoryCount SIDs are in the users SIDHistory"
        write-verbose "    - $TotalGroupSIDHistoryCount SIDs are in the users group SIDHistory"
        write-verbose "- The current userAccountControl value is $UACValue"
        If ($Enabled) {
          write-verbose "  - The account is enabled"
        } Else {
          write-verbose "  - The account is disabled"
        }
        If ($TrustedforDelegation -eq $false) {
          write-verbose "  - The account is not trusted for delegation"
        } Else {
          write-verbose "  - The account is trusted for delegation"
        }
        write-verbose "- The primary group is $primarygroupname"
      }

      # Calculate the current token size, taking into account whether or not the account is trusted for delegation or not.
      $TokenSize = 1200 + (40 * ($SecurityDomainLocalScope + $SecurityGlobalExternalScope + $SecurityUniversalExternalScope + $UserSIDHistoryCount + $TotalGroupSIDHistoryCount)) + (8 * ($SecurityGlobalInternalScope  + $SecurityUniversalInternalScope))
      if ($TrustedforDelegation) {
        $TokenSize = 2 * $TokenSize
      }
      write-verbose "- Therefore the estimated Token size is $Tokensize"

      If ($ProcessSingleAccount) {
        write-verbose "- Refer to $SingleAccountOutputFile for a detailed output"
      }

      If ($TokenSize -le 12000) {
        $TotalTokensLessThan8K ++
        If ($TokenSize -gt 8192) {
          $TotalTokensBetween8and12K ++
        }
      } elseIf ($TokenSize -le 48000) {
        $TotalTokensGreaterThan12K ++
      } else {
        $TotalTokensGreaterThan48K ++
      }

      If ($TokenSize -gt $LargestTokenSize) {
        $LargestTokenSize = $TokenSize
        $LargestTokenUser = $SamAccountName
      }

      If ($TokenSize -ge $TokensSizeThreshold) {
        $obj = New-Object -TypeName PSObject
        $obj | Add-Member -MemberType NoteProperty -Name "Domain" -value $domain
        $obj | Add-Member -MemberType NoteProperty -Name "SamAccountName" -value $SamAccountName
        $obj | Add-Member -MemberType NoteProperty -Name "TokenSize" -value $TokenSize
        $obj | Add-Member -MemberType NoteProperty -Name "Memberships" -value $GroupCount
        $obj | Add-Member -MemberType NoteProperty -Name "DomainLocal" -value $SecurityDomainLocalScope
        $obj | Add-Member -MemberType NoteProperty -Name "GlobalInternal" -value $SecurityGlobalInternalScope
        $obj | Add-Member -MemberType NoteProperty -Name "GlobalExternal" -value $SecurityGlobalExternalScope
        $obj | Add-Member -MemberType NoteProperty -Name "UniversalInternal" -value $SecurityUniversalInternalScope
        $obj | Add-Member -MemberType NoteProperty -Name "UniversalExternal" -value $SecurityUniversalExternalScope
        $obj | Add-Member -MemberType NoteProperty -Name "UserSIDHistory" -value $UserSIDHistoryCount
        $obj | Add-Member -MemberType NoteProperty -Name "GroupSIDHistory" -value $TotalGroupSIDHistoryCount
        $obj | Add-Member -MemberType NoteProperty -Name "UACValue" -value $UACValue
        $obj | Add-Member -MemberType NoteProperty -Name "Enabled" -value $Enabled
        $obj | Add-Member -MemberType NoteProperty -Name "TrustedforDelegation" -value $TrustedforDelegation
        $obj | Add-Member -MemberType NoteProperty -Name "LastLogon" -value $lastLogon
        $arrayoftopusers += $obj

        # PowerShell V2 doesn't have an Append parameter for the Export-Csv cmdlet. Out-File does, but it's
        # very difficult to get the formatting right, especially if you want to use quotes around each item
        # and add a delimeter. However, we can easily do this by piping the object using the ConvertTo-Csv,
        # Select-Object and Out-File cmdlets instead.
        if ($PSVersionTable.PSVersion.Major -gt 2) {
          $obj | Export-Csv -Path "$ReferenceFile" -Append -Delimiter $Delimiter -NoTypeInformation -Encoding ASCII
        } Else {
          if (!(Test-Path -path $ReferenceFile)) {
            $obj | ConvertTo-Csv -NoTypeInformation -Delimiter $Delimiter | Select-Object -First 1 | Out-File -Encoding ascii -filepath "$ReferenceFile"
          }
          $obj | ConvertTo-Csv -NoTypeInformation -Delimiter $Delimiter | Select-Object -Skip 1 | Out-File -Encoding ascii -filepath "$ReferenceFile" -append -noclobber
        }

        If ($ProcessSingleAccount -eq $False) {
          # Manage an array of the top X users as per the $TopUsers variable.
          $arrayoftopusers | Sort-Object TokenSize -descending | select-object -first $TopUsers | out-null
        }

      }

    } Else {
      write-verbose "Excluding user $SamAccountName in domain $domain"
      $TotalExcludedUsers ++
    }

    If ($ConsoleOutput -AND $ProcessSingleAccount -eq $False) {
      $percent = "{0:P}" -f ($TotalUsersProcessed/$UserCount)
      write-verbose "- Processed $TotalUsersProcessed of $UserCount user accounts = $percent complete."
      # Add a blank line
      Write-Host " "
    }

    $garbageCounter++
    if ($garbageCounter -eq 500) {
      [System.GC]::Collect()
      $garbageCounter = 0
    }

  }
  # Dispose of the search and results properly to avoid a memory leak
  $ADSearcher.Dispose() | Out-Null
  [System.GC]::Collect() | Out-Null

  If ($OutputSummary) {
    write-verbose " Summary:"
    write-verbose " - Processed $UserCount user accounts."
    If ($TotalExcludedUsers -gt 0) {
      write-verbose " - Excluded $TotalExcludedUsers user accounts."
    }
    write-verbose " - $TotalTokensLessThan8K have a calculated token size of less than or equal to 12000 bytes."
    If ($TotalTokensLessThan8K -gt 0) {
      write-verbose "  - These users are good."
    }
    If ($TotalTokensBetween8and12K -gt 0) {
      write-verbose "  - Although $TotalTokensBetween8and12K of these user accounts have tokens above 8K and should therefore be reviewed."
    }
    write-verbose " - $TotalTokensGreaterThan12K have a calculated token size larger than 12000 bytes."
    If ($TotalTokensGreaterThan12K -gt 0) {
      write-verbose "   - These users will be okay if you have increased the MaxTokenSize to 48000 bytes across the domain/forest."
      write-verbose "   - Consider reducing direct and transitive (nested) group memberships."
    }
    If ($TotalTokensGreaterThan48K -gt 0) {
      write-warning " - $TotalTokensGreaterThan48K have a calculated token size larger than 48000 bytes."
      write-warning "   - These users will have problems. Do NOT increase the MaxTokenSize beyond 48000 bytes."
      write-warning "   - Reduce the direct and transitive (nested) group memberships."
    }
    write-verbose " - $LargestTokenUser has the largest calculated token size of $LargestTokenSize bytes in the $domain domain."
  }

  If ($ProcessSingleAccount -eq $False) {
    # Write the $arrayoftopusers to a CSV.
    # $arrayoftopusers | export-csv -notype -path "$ReferenceFileTopUsers" -Delimiter $Delimiter
    if ($PSVersionTable.PSVersion.Major -gt 2) {
      $arrayoftopusers | Export-Csv -Path "$ReferenceFileTopUsers" -Append -Delimiter $Delimiter -NoTypeInformation -Encoding ASCII
    } Else {
      if (!(Test-Path -path $ReferenceFileTopUsers)) {
        $arrayoftopusers | ConvertTo-Csv -NoTypeInformation -Delimiter $Delimiter | Select-Object -First 1 | Out-File -Encoding ascii -filepath "$ReferenceFileTopUsers"
      }
      $arrayoftopusers | ConvertTo-Csv -NoTypeInformation -Delimiter $Delimiter | Select-Object -Skip 1 | Out-File -Encoding ascii -filepath "$ReferenceFileTopUsers" -append -noclobber
    }
  }

}
