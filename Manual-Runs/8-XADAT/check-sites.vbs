On Error Resume Next

Set objRootDSE = GetObject("LDAP://RootDSE")
strConfigurationNC = objRootDSE.Get("configurationNamingContext")
 
strSitesContainer = "LDAP://cn=Sites," & strConfigurationNC
Set objSitesContainer = GetObject(strSitesContainer)
objSitesContainer.Filter = Array("site")

strSubnetsContainer = "LDAP://cn=Subnets,cn=Sites," & strConfigurationNC
Set objSubnetsContainer = GetObject(strSubnetsContainer)
objSubnetsContainer.Filter = Array("subnet")
Set objHash = CreateObject("Scripting.Dictionary")

x = 0
y = 0

For Each objSite In objSitesContainer

    Redim Preserve SitesAD (x + 1)

    SitesAD(x) = objSite.CN
    x = x + 1
Next

For Each objSubnet In objSubnetsContainer
    objSubnet.GetInfoEx Array("siteObject"), 0
    strSiteObjectDN = objSubnet.Get("siteObject")
    strSiteObjectName = Split(Split(strSiteObjectDN, ",")(0), "=")(1)

    If objHash.Exists(strSiteObjectName) Then
        objHash(strSiteObjectName) = objHash(strSiteObjectName) & "," & _
            Split(objSubnet.Name, "=")(1)

    Else
        objHash.Add strSiteObjectName, Split(objSubnet.Name, "=")(1)
    End If

  For Each strKey In objHash.Keys

     Redim Preserve subSites(y + 1)
     subSites(y) = strkey
     y = y + 1
  Next

Next

For i = 0 to (Ubound(SitesAD) - 1)
   For j = 0 to (Ubound(subSites) - 1)
      if SitesAD(i) = subSites(j) then
         SitesAD(i) = "xxxx"
      end if
   Next
Next

cnt = 0
For k = 0 to (Ubound(SitesAD) - 1)
    if SitesAD(k) <> "xxxx" then
       wscript.echo "AD Site: " & SitesAD(k) & " has no subnets assigned to it"
       cnt = cnt + 1

    end if
Next

if cnt = 0 then
   wscript.echo "No issues found with AD sites.."
end if