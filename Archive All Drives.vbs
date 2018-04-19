Option Explicit
Dim oFS, oDrive, oFileTS, sOutPut, oRoot
Set oFS = CreateObject("Scripting.FileSystemObject")
Set oFileTS = oFS.CreateTextFile("c:\ArchiveData.csv")
oFileTS.WriteLine "File Name" & "," & "Last Accessed"
For Each oDrive In oFS.Drives
    If oDrive.DriveType = 2 Then
     	Set oRoot = oDrive.RootFolder
     	WriteFile(oRoot)
	End If 
Next
WScript.Echo "Finished."	

Sub WriteFile(ByRef oFol)
	Dim oFolder, oFiles
	On Error Resume Next
	For Each oFolder in oFol.SubFolders
		If Err.Number = 0 Then
			For Each oFiles in oFolder.Files
				If Err.Number = 0 Then oFileTS.WriteLine oFiles.Path & "," & oFiles.DateLastAccessed
			Next
	    End If
	    Err.Clear
	    WriteFile(oFolder)
	Next	
End Sub
