Const mask_8 = 16777214
Const mask_9 = 8388606
Const mask_10 = 4194302
Const mask_11 = 2097150
Const mask_12 = 1048574
Const mask_13 = 524286
Const mask_14 = 262142
Const mask_15 = 131070
Const mask_16= 65534
Const mask_17 = 32766
Const mask_18 =16382
Const mask_19 = 8190
Const mask_20 = 4094
Const mask_21 = 2046
Const mask_22 = 1022
Const mask_23 = 510
Const mask_24 = 254
Const mask_25 = 126
Const mask_26 = 62
Const mask_27= 30
Const mask_28 = 14
Const mask_29 = 6
Const mask_30 = 2
Const mask_31 = 1
Const mask_32 = 0

Set oRootDSE = GetObject("LDAP://RootDSE")
Set oSysInfo = CreateObject("ADSystemInfo")
sDNSDom = UCase(oRootDSE.Get("RootDomainNamingContext"))
sFQDN = ucase(oSysInfo.ForestDNSName)


Set oConn = CreateObject("ADODB.Connection")
oConn.Provider = "ADsDSOObject"
oConn.Open "Active Directory Provider"

Set oCmd = CreateObject("ADODB.Command")
oCmd.ActiveConnection = oConn

sADsPath = "<LDAP://CN=Subnets,CN=Sites,CN=Configuration," & sDNSDom & ">"
oCmd.Properties("Page Size") = 100
oCmd.CommandText = sADsPath & ";(objectCategory=subnet);ADsPath,cn,whenCreated,siteObject;subTree"
Set oRecSet = oCmd.Execute

'
' Set the loop counter 
'

i = 0
k = 0
  While Not oRecSet.EOF
  
       ReDim Preserve sAddr (i + 1)
       ReDim Preserve sLower (i + 1)
       ReDim Preserve sUpper (i + 1)
       ReDim Preserve sLoc (i + 1)
       ReDim Preserve sDate (i + 1)

       sDate(i) = oRecSet.Fields("whenCreated").Value
       sDate(i) = Left(sDate(i), InStr(1, sDate(i), " ", vbTextCompare) - 1)

       sAddr(i) = oRecSet.Fields("cn").Value
       sLoc(i)  = Left(oRecSet.Fields("siteObject").Value, (Instr(oRecSet.Fields("siteObject").Value, ",")-1))
         rawAddr = Split (sAddr(i), "/")
         cleanAddr = rawAddr(0)

         sOctal = Split (cleanAddr, ".")
         sOctal(0) = 256^3 * sOctal(0)
         sOctal(1) = 256^2 * sOctal(1)
         sOctal(2) = 256 * sOctal(2)
         sOctalsum = sOctal(0) + sOctal(1) + sOctal(2) + sOctal(3)


         If Trim(sLoc(i)) = "" then

           wscript.echo "Subnet: " & rawAddr(0) & " is an orphaned subnet and was created on " & sDate(i)

           k = k + 1

         End if



           sLower(i) = sOctalsum + 1


         select case rawAddr(1)

             case 8
                  sUpper(i) = sLower(i) + mask_8
             case 9
                  sUpper(i) = sLower(i) + mask_9
             case 10
                  sUpper(i) = sLower(i) + mask_10
             case 11
                  sUpper(i) = sLower(i) + mask_11
             case 12
                  sUpper(i) = sLower(i) + mask_12
             case 13
                  sUpper(i) = sLower(i) + mask_13
             case 14
                  sUpper(i) = sLower(i) + mask_14
             case 15
                  sUpper(i) = sLower(i) + mask_15
             case 16
                  sUpper(i) = sLower(i) + mask_16
             case 17
                  sUpper(i) = sLower(i) + mask_17
             case 18
                  sUpper(i) = sLower(i) + mask_18
             case 19
                  sUpper(i) = sLower(i) + mask_19
             case 20
                  sUpper(i) = sLower(i) + mask_20
             case 21
                  sUpper(i) = sLower(i) + mask_21
             case 22
                  sUpper(i) = sLower(i) + mask_22
             case 23
                  sUpper(i) = sLower(i) + mask_23
             case 24
                  sUpper(i) = sLower(i) + mask_24
             case 25
                  sUpper(i) = sLower(i) + mask_25
             case 26
                  sUpper(i) = sLower(i) + mask_26
             case 27
                  sUpper(i) = sLower(i) + mask_27
             case 28
                  sUpper(i) = sLower(i) + mask_28
             case 29
                  sUpper(i) = sLower(i) + mask_29
             case 30
                  sUpper(i) = sLower(i) + mask_30
             case 31
                  sUpper(i) = sLower(i) + mask_31
             case 32
                  sUpper(i) = sLower(i) + mask_32

             case else
                  sUpper(i) = sUpper(i)
         End Select
     

        i = i + 1
        oRecSet.MoveNext
    Wend
   wscript.echo ""

   wscript.echo "The total number of subnets is: " & i
   wscript.echo ""
   wscript.echo ""


' Find and report duplicates
'
'
cnt = 0
bad = 0
F1 = False
F2 = False

While not (cnt > uBound(sLower))
  For m = 0 to ubound(sLower)

     if not (m = cnt) then
       F1 = sLower(cnt) >= sLower(m) AND sLower(cnt) <= sUpper(m)
       F2 = sUpper(cnt) >= sLower(m) AND sUpper(cnt) <= sUpper(m)

     End if

     if F1 or F2 Then
       wscript.echo "The subnet: " & sAddr(cnt) & " from Location " & sLoc(cnt) & " has a conflict with subnet: " & sAddr(m) & " in Location " & sLoc(m)
       bad = bad + 1
     end if

   F1 = False
   F2 = False
  Next
 cnt = cnt + 1

Wend 

wscript.echo ""
If bad > o then

wscript.echo "The number of conflicting subnets is: " & bad
Else

wscript.echo "There are no conflicting subnets in the forest."
End if

wscript.echo ""

If k = 0 then
wscript.echo "There are no orphaned subnets in the forest."
End if
