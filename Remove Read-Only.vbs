Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("C:\Documents and Settings\Administrator.CLC-9A02C69D6D8\Desktop\2274C Slides")
Set colFiles = objFolder.Files
For Each objFile in colFiles
    If objFile.Attributes AND 1 Then
        objFile.Attributes = objFile.Attributes XOR 1
    End If
Next

