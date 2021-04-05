Set objRoot = GetObject("LDAP://rootDSE")
Set objConfig = GetObject("LDAP://" & objRoot.Get("ConfigurationNamingContext"))
objConfig.Filter = Array("crossRefContainer")

strCount = 1

strRoot = objRoot.Get("RootDomainNamingContext")

Wscript.Echo strCount & "*" & Chr(34) & strRoot& Chr(34)

For Each objPartition in objConfig
    strPartition = "LDAP://" & objPartition.Get("distinguishedName")

Next

strcount = strCount + 1

Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open "Provider=ADsDSOObject;"

Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000
    
objCommand.CommandText = "<" & strPartition & ">;(&(systemFlags=3));nCName,systemFlags;subTree"
Set objRecordset = objCommand.Execute
    
' List all domains in Forest
'
Do While Not objRecordset.EOF
      if not strRoot = (objRecordset.Fields(0)) then
	  wscript.echo strcount & "*" & Chr(34) & (objRecordset.Fields(0))& Chr(34)
          strcount = strCount + 1
      end if

       objRecordset.MoveNext
Loop
     
