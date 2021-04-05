Const adVarChar = 200
Const MaxCharacters = 255
Const ForReading = 1
Const ForWriting = 2
Dim ArgObj, var1

Set ArgObj = WScript.Arguments 
var1 = ArgObj(0)

Set DataList = CreateObject("ADOR.Recordset")
DataList.Fields.Append "GPOGUID", adVarChar, MaxCharacters
DataList.Open

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(var1, ForReading)

Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
    strLine = Trim(strLine)
    If Len(strLine) > 0 Then
      DataList.AddNew
      DataList("GPOGUID") = strLine
      DataList.Update
    End if
Loop

objFile.Close

DataList.Sort = "GPOGUID"

DataList.MoveFirst
Do Until DataList.EOF
    strText = strText & DataList.Fields.Item("GPOGUID") & vbCrLf
    DataList.MoveNext

Loop

Set objFile = objFSO.OpenTextFile(var1, ForWriting)

objFile.WriteLine strText
objFile.Close
