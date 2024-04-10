<%

' File                : getFilesList.asp
' Project             : PROG2001 - Web Design and Development
' Programmer          : Minchul Hwang - 8818858
' First version       : Dec. 2. 2023
' Description         : This file is connected to startPage.asp and is responsible for loading files in the folder.


Dim CheckFile
Set CheckFile = Server.CreateObject("Scripting.FileSystemObject")

Dim folderPath
folderPath = Server.MapPath("MyFiles\")

Dim folder
Set folder = CheckFile.GetFolder(folderPath)

Dim fileNames
fileNames = Array()

'Get files from folder
For Each file In folder.Files
    ReDim Preserve fileNames(UBound(fileNames) + 1)
    fileNames(UBound(fileNames)) = CheckFile.GetBaseName(file.Name)
Next

Response.Write JsonArray(fileNames)

Set CheckFile = Nothing

' set json object from a file 
Function JsonArray(arr)
    Dim result
    result = "["
    For i = 0 To UBound(arr)
        If i > 0 Then result = result & ","
        result = result & """" & arr(i) & """"
    Next
    result = result & "]"
    JsonArray = result
End Function
%>