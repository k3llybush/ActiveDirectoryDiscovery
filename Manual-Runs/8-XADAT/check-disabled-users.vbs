On Error Resume Next

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

Set objDomain = GetObject("LDAP://RootDSE") 
strDomain = objDomain.Get("DefaultNamingContext")

objCommand.Properties("Page Size") = 1000

objCommand.CommandText = _
    "<LDAP://" & strDomain & ">;(&(objectCategory=User)" & _
        "(userAccountControl:1.2.840.113556.1.4.803:=2));displayname,saMAccountName;Subtree"  
Set objRecordSet = objCommand.Execute

wscript.echo "Disabled Users"
wscript.echo "-------------------------------------------------------------------------------"

objRecordSet.MoveFirst
cnt = 0

Do Until objRecordSet.EOF
    Wscript.Echo objRecordSet.Fields("displayname").Value & vbTab & objRecordSet.Fields("saMAccountName").Value
    cnt = cnt + 1
    objRecordSet.MoveNext
Loop

if cnt = 0 then
  wscript.echo "There are no disabled users.."
end if