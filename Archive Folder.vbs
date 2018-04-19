Option Explicit
Dim oFS, oDrive, oFileTS, sOutPut, oRoot, dDate, oSh, oFiles
Dim oShell, oRootFolder, oPath, objFolderItem, aFinished
Const WINDOW_HANDLE = 0
Const NO_OPTIONS = 0
dDate = InputBox("Create List of Files Last Accessed Before [date]:" _
        & vbCrLf & vbCrLf & "(Default is files that have not been " _
        & "accessed for the last 1 year.", _
        "Archive Files", (Date - 365))
If dDate = "" Then WScript.Quit
Set oShell = CreateObject("Shell.Application")      
On Error Resume Next
Set oRootFolder = oShell.BrowseForFolder _
    (WINDOW_HANDLE, "Select a folder to scan:", NO_OPTIONS)
If Err.Number <> 0 Then WScript.Quit
If oRootFolder = Null Then WScript.Quit       
Set objFolderItem = oRootFolder.Self
oPath = objFolderItem.Path

Set oFS = CreateObject("Scripting.FileSystemObject")
Set oFileTS = oFS.CreateTextFile("c:\ArchiveData.csv")
Set oRoot = oFS.GetFolder(oPath)
oFileTS.WriteLine "File Name" & "," & "Last Accessed"
For Each oFiles in oRoot.Files
    If Err.Number = 0 Then
        If oFiles.DateLastAccessed <= CDate(dDate) Then
			oFileTS.WriteLine _
				oFiles.Path & "," & oFiles.DateLastAccessed
		End If
	End If
Next
WriteFile(oRoot)
aFinished = MsgBox("Finished. The results can be found in C:\ArchiveData.csv" _
    & vbCrLf & vbCrLf & "Do you want to open this file?", vbYesNo)
If aFinished = vbYes Then
    set oSh = CreateObject("WScript.Shell")
    oSh.Run "Excel.exe C:\ArchiveData.csv"
Else WScript.Quit
End If

Sub WriteFile(oFol)
	Dim oFolder, oFiles
	On Error Resume Next
	For Each oFolder in oFol.SubFolders
		If Err.Number = 0 Then
			For Each oFiles in oFolder.Files
				If Err.Number = 0 Then
				    If oFiles.DateLastAccessed <= CDate(dDate) Then
				        oFileTS.WriteLine _
				            oFiles.Path & "," & oFiles.DateLastAccessed
				    End If
				End If
			Next
	    End If
	    Err.Clear
	    WriteFile(oFolder)
	Next	
End Sub
