If Wscript.Arguments.Count = 0 Then
  Wscript.Echo "Usage: [Win_2003/2008_DC_Name]"
  Wscript.Quit(1)
Else
  strComputer = Trim(Wscript.Arguments(0))
End If
 
strComputer = UCase(strComputer)
 
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!"_
   & strComputer & "\root\MicrosoftActiveDirectory")

cnt = 0
 
Set colTrustList = objWMIService.ExecQuery _
    ("Select * from Microsoft_DomainTrustStatus")
 
For each objTrust in colTrustList
    Wscript.Echo " ******** "
    Wscript.Echo "Trusted domain: " & objTrust.TrustedDomain
    Wscript.Echo "Trusted DC name: " & objTrust.TrustedDCName
    Wscript.Echo "Trust status (0 is good): " & objTrust.TrustStatus
    Wscript.Echo "Is Trust Ok: " & objTrust.TrustIsOK

    cnt = cnt + 1
Next
 
if cnt = 0 then
   wscript.echo "There are no trusts setup with this domain...."
end if