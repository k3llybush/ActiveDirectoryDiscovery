
'Declare Variables:
Dim oRootDSE, oCon, oCom, oRS, oDic, oKey
Dim sDNC, sGPOQuery, sOUQuery

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
While Not oRS.EOF
	oDic.Add oRS.Fields("cn").Value, oRS.Fields("displayName").Value
	oRS.MoveNext
Wend
Set oRS = Nothing

'Retrieve OUs:
oCom.CommandText = sOUQuery
Set oRS = oCom.Execute
While Not oRS.EOF
	Wscript.Echo oRS.Fields("distinguishedName").Value
	For Each oKey In oDic.Keys
		If InStr(oRS.Fields("gPLink").Value, oKey) > 0 Then
			Wscript.Echo vbTab & "GPO Link: " & oDic(oKey)
		End If
	Next
	Wscript.Echo ""
	oRS.MoveNext
Wend

'Clean Objects:
Set oRS = Nothing
Set oCom = Nothing
oCon.Close
Set oCon = Nothing