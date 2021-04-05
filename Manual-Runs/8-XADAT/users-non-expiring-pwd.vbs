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
    "<LDAP://" & strDomain & ">;" & _
        "(&(objectCategory=User)(userAccountControl:1.2.840.113556.1.4.803:=65536));" & _
            "Name,saMAccountName;Subtree"
Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst
wscript.echo "Name           Logon ID"
wscript.echo "--------------------------------------------------------------------------"

cnt = 0

Do Until objRecordSet.EOF
    Wscript.Echo objRecordSet.Fields("Name").Value & vbTab & objRecordSet.Fields("saMAccountName").Value 
    cnt = cnt + 1
    objRecordSet.MoveNext
Loop

if cnt = 0 then
  wcript.echo "There are no accounts setup with non-expiring passwords.."
end if