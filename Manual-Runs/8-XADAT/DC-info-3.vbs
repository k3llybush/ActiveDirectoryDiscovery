
On Error Resume Next

Set objDomain = GetObject("LDAP://RootDSE") 
strDomain = objDomain.Get("DefaultNamingContext")
Set objOU = GetObject("LDAP://ou=Domain Controllers," & strDomain)


For Each strServer In objOU


   If Ping (strServer.Cn) then

'  set shell2 = CreateObject("Wscript.Shell")
'  set exec2 = shell2.Exec("psinfo -d \\" & strServer.Cn)
'   wscript.echo exec2.Stdout.ReadAll


'  set shell3 = CreateObject("Wscript.Shell")
'  set exec3 = shell3.Exec("dcdiag /test:netlogons /s:" & strServer.Cn)
'   wscript.echo exec3.Stdout.ReadAll

'  set exec3 = shell3.Exec("dcdiag /test:Replications /s:" & strServer.Cn)
'   wscript.echo exec3.Stdout.ReadAll


  set shell4 = CreateObject("Wscript.Shell")
  set exec4 = shell4.Exec("reg query \\" & strServer.Cn & "\HKLM\System\CurrentControlSet\services\W32Time\Parameters /v Type")

       if Instr(UCase(exec4.Stdout.ReadAll), "NT5DS") then
         wscript.echo strServer.cn & " has NT5DS as the Time Sync configuration"
       end if
       
       if Instr(UCase(exec4.Stdout.ReadAll), "NTP") then
         wscript.echo strServer.cn & " has NTP as the Time Sync configuration"
       end if

       if Instr(UCase(exec4.Stdout.ReadAll), "") then
         wscript.echo strServer.cn & " Was inaccessible"
       end if

  End if

Next

'------------------------------------------------------
 Function Ping(Target)
 Dim Results

   Set shell = CreateObject("Wscript.Shell")
   Set exec = shell.Exec("ping -n 1 -w 2000 " & Target)
   results = LCase(exec.StdOut.ReadAll)

   Ping = (InStr(results, "reply from") > 0)

 End Function
'------------------------------------------------------