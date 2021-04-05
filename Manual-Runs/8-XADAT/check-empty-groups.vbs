'Define Constants
Const ADS_SCOPE_SUBTREE = 2 ' Search target object and all sub levels
 
'Set Variables
DQ = Chr(34) 'Double Quote
 
'Create Objects
Set objShell = CreateObject("Wscript.Shell")
 

'Construct an ADsPath to the Current Domain with rootDSE
Set objRootDSE = GetObject("LDAP://rootDSE")
strADsPath = "LDAP://" & objRootDSE.Get("defaultNamingContext")
 
'Connect to Active Directory
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
 
objCommand.CommandText = "SELECT ADsPath FROM '" & strADsPath & _
"'" & " WHERE objectCategory='group' AND NOT member='*'"


Set objRecordSet = objCommand.Execute
 
If objRecordSet.EOF Then
	WScript.echo "There are no empty groups found.."
	WScript.quit
Else
	WScript.Echo "List of empty groups"
	WScript.Echo "============================================================="
	objRecordSet.MoveFirst
	Do Until objRecordSet.EOF
          strGroupName = objRecordSet.Fields("ADsPath").Value
          WScript.Echo Right(strGroupName, (len(strGroupname) - (instr(strGroupName, "/")+ 1)))
          objRecordSet.MoveNext
        Loop
End If
 
