'Declare Variables:
Dim oRootDSE, oCon, oCom, oRS, oDic
Dim sDNC, sGPOQuery

'Retrieve LDAP Naming Context:
Set oRootDSE = GetObject("LDAP://rootDSE")
sDNC = "LDAP://" & oRootDSE.Get("defaultNamingContext")
Set oRootDSE = Nothing

'Define LDAP Queries:
sGPOQuery = "<" & sDNC & ">;(objectClass=groupPolicyContainer);name;subtree"

'Connect to AD with ADODB:
Set oCon = CreateObject("ADODB.Connection")
Set oCom = CreateObject("ADODB.Command")
oCon.Provider = "ADsDSOObject"
oCon.Open "Active Directory Provider"
Set oCom.ActiveConnection = oCon
oCom.Properties("Page Size") = 1000
Set oDic = CreateObject("Scripting.Dictionary")

'Retrieve GPO GUIDs:
oCom.CommandText = sGPOQuery
Set oRS = oCom.Execute
While Not oRS.EOF
'        oDic.Add oRS.Fields("name").Value
	wscript.echo oRS.Fields("name").Value
	oRS.MoveNext
Wend
Set oRS = Nothing

'Clean Objects:
Set oRS = Nothing
Set oCom = Nothing
oCon.Close
Set oCon = Nothing