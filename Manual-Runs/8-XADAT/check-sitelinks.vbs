On Error Resume Next

RowCounter = 0

set objRootDSE = GetObject("LDAP://RootDSE")
set objSitesCont = GetObject("LDAP://cn=sites," & _
                             objRootDSE.Get("configurationNamingContext") )
objSitesCont.Filter = Array("site")
for each objSite in objSitesCont
   
  
   
   strSite = objSite.Get("cn")
   
   set objServersCont = GetObject("LDAP://cn=Servers,cn=" & strSite & ",cn=Sites," & objRootDSE.Get("configurationNamingContext") )
   objServersCont.Filter = Array("server")

   for each objServer in objServersCont      
      strServer = objServer.Get("cn")

   next
   
   subnets = objSite.GetEx("siteObjectBL")

   for each subnet in subnets      
      subnetArray = split(subnet, ",")
      strSubnet = Mid(subnetArray(0), 4)
      ThirdColRow = ThirdColRow + 1
   next
   
   strSiteDN = "cn=" & strSite & ",cn=sites," & _
               objRootDSE.Get("ConfigurationNamingContext")

   strBase    =  "<LDAP://cn=Inter-site Transports,cn=sites," _
              & objRootDSE.Get("ConfigurationNamingContext") & ">;"
   strFilter  = "(&(objectcategory=siteLink)" & _
             "(siteList=" & strSiteDN & "));" 
   strAttrs   = "name;"
   strScope   = "subtree"

   set objConn = CreateObject("ADODB.Connection")
   objConn.Provider = "ADsDSOObject"
   objConn.Open "Active Directory Provider"
   set objRS = objConn.Execute(strBase & strFilter & strAttrs & strScope)
   
   if objRS.RecordCount > 0 then
      objRS.MoveFirst

      while Not objRS.EOF
         strLink = objRS.Fields(0).Value
         set objSiteLink = GetObject("LDAP://cn=" & strLink & ",cn=IP,cn=Inter-site Transports,cn=sites," & objRootDSE.Get("configurationNamingContext") )
          FifthColRow = FourthColRow
         sitesIn = objSiteLink.GetEx("siteList")
         cnt = 0
         for each siteIn in sitesIn
            siteArray = split(siteIn, ",")
            strsiteIn = Mid(siteArray(0), 4)
           cnt = cnt + 1
           if cnt = 3 then
              wscript.echo "The site link " & strLink &  " has more than two sites in it!!"
              RowCounter = RowCounter + 1

           end if
         next
           if cnt = 0 then
              wscript.echo "The site link " & strLink &  " has no sites in it!!"
              RowCounter = RowCounter + 1

           end if
           if cnt = 1 then
              wscript.echo "The site link " & strLink &  " has only one site in it!!"
              RowCounter = RowCounter + 1

           end if

         objRS.MoveNext
      wend

   end if
   
next

If RowCounter = 0 then
   wscript.echo "There are no known issues with Site Links..."
end if
