set oFSO = CreateObject("Scripting.FileSystemObject")
spaces = ""
EnumSubFolders oFSO.GetFolder("C:\"), spaces
Function EnumSubFolders(Folder, ByVal spaces)
   On Error Resume Next
   For Each oSubFolder in Folder.SubFolders
       strOutput = spaces & "+" & oSubFolder.Name
       If Err.Number = 0 Then WScript.Echo strOutput
       Err.Clear
       EnumSubFolders oSubFolder, spaces & "  "
   Next
End Function