'Retrieve LDAP Naming Context:
Set oRootDSE = GetObject("LDAP://rootDSE")
sDNC = "LDAP://" & oRootDSE.Get("defaultNamingContext")
Set oRootDSE = Nothing

'Define LDAP Queries:
sGPOQuery = "<" & sDNC & ">;(objectClass=groupPolicyContainer);cn,displayName;subtree"
sOUQuery = "<" & sDNC & ">;(|(objectClass=domainDNS)(objectClass=organizationalUnit));cn,distinguishedName,gPLink;subtree"

'Connect to AD with ADODB:
Set oCon = CreateObject("ADODB.Connection")
Set oCom = CreateObject("ADODB.Command")
oCon.Provider = "ADsDSOObject"
oCon.Open "Active Directory Provider"
Set oCom.ActiveConnection = oCon
oCom.Properties("Page Size") = 1000
 Set oDic = CreateObject("Scripting.Dictionary")

'Retrieve GPOs:
oCom.CommandText = sGPOQuery
Set oRS = oCom.Execute

x = 0
While Not oRS.EOF
	oDic.Add oRS.Fields("cn").Value, oRS.Fields("displayName").Value
Redim Preserve strGPOs (x + 1)

   strGPOs(x) = oRS.Fields("displayName").Value

  oRS.MoveNext

x = x + 1
Wend

Set oRS = Nothing

'Retrieve OUs:
oCom.CommandText = sOUQuery
Set oRS = oCom.Execute

While Not oRS.EOF
	' Wscript.Echo oRS.Fields("distinguishedName").Value
	For Each oKey In oDic.Keys
		If InStr(oRS.Fields("gPLink").Value, oKey) > 0 Then
'			Wscript.Echo oDic(oKey)

                         For G = 0 to (ubound(strGPOs) - 1)

                         if strGPOs(G) = oDic(oKey) then

                             strGPOs(G) = "xxxx"
                         end if

                         next
                        
		End If
	Next
	oRS.MoveNext
Wend

cnt = 0

For n = lbound(strGPOs) to (ubound(strGPOs) - 1)
 if strGPOs(n) <> "xxxx" then
  Wscript.echo "GPO name: " & strGPOs(n)
  cnt = cnt + 1

 End if
next

if cnt = 0 then
 wscript.echo "No orphaned GPOs found.."
end if

'Clean Objects:
Set oRS = Nothing
Set oCom = Nothing
oCon.Close
Set oCon = Nothing