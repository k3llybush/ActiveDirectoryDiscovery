' Get Forest's root
'
Set objRoot = GetObject("LDAP://rootDSE")

' Get root's Configuration
'
Set objConfig = GetObject("LDAP://" & objRoot.Get("ConfigurationNamingContext"))

' Search for the Partitions container in root's Configuration
'
objConfig.Filter = Array("crossRefContainer")
For Each objPartition in objConfig
    strPartition = "LDAP://" & objPartition.Get("distinguishedName")
Next

Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open "Provider=ADsDSOObject;"

Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000
    
objCommand.CommandText = "<" & strPartition & ">;(&(systemFlags=3));nCName,systemFlags;subTree"
Set objRecordset = objCommand.Execute
    
cnt = 0

Do While Not objRecordset.EOF
	' List all DCs in one domain
	'
	Set objDCs = GetObject("LDAP://OU=Domain Controllers," & objRecordset.Fields(0))	
	For Each objDC in objDCs
         if not (objDC.Get("userAccountcontrol")) = 532480 then

		wscript.echo "Domain Controller: " & (objDC.Get("Name")) & _
                 " has incorrect UAC setting.  It should be 532480 and it is: " & _
		   (objDC.Get("userAccountcontrol"))
                cnt = cnt + 1

	  End if
	
	Next
	
	objRecordset.MoveNext
Loop
     
if cnt = 0 then
  wscript.echo "The domain controller UAC settings are configured correctly.."
end if
