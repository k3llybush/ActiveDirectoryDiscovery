Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objAD = objFSO.OpenTextFile("GPOs-in-AD-but-not-in-sysvol.txt", 1)
Set objSYS = objFSO.OpenTextFile("GPOs-in-SYSVOL-but-not-in-AD.txt", 1)


'Now, read the contents of the file into a string
Dim s, t
s = objAD.ReadAll
t = objSYS.ReadAll

Dim aFile, sFile

aFile = split(s, vbCrLf)
sFile = split(t, vbCrLf)

 For i = 0 to (Ubound(sFile) - 1)
   For j = 0 to (Ubound(aFile) - 1)

       if sFile(i) = aFile(j) then
          aFile(j) = "xxxx"
          sFile(i) = "xxxx"

       end if
   Next
 Next

ADcnt = 0
SYScnt = 0

wscript.echo "GPOs that are in AD, but not in SYSVOL"
wscript.echo "**************************************************"
wscript.echo ""

For k = 0 to (Ubound(aFile) - 1)
    if aFile(k) <> "xxxx" then
       ADcnt = ADcnt + 1
       wscript.echo aFile(k)
    end if
Next
    if ADcnt = 0 then
       wscript.echo "There are no orphaned GPOs in AD"
    end if

wscript.echo ""
wscript.echo ""

wscript.echo "GPOs that are in SYSVOL, but not in AD"
wscript.echo "**************************************************"
wscript.echo ""

For l = 0 to (Ubound(sFile) - 1)
    if sFile(l) <> "xxxx" then
       SYScnt = SYScnt + 1
       wscript.echo sFile(l)
    end if
Next
    if SYScnt = 0 then
       wscript.echo "There are no orphaned GPOs in SYSVOL"
    end if

 
      