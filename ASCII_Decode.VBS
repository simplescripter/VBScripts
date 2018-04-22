'-------------------------------------------------------
'
'  Shawn Stugart
'=======================================================
Option Explicit

Dim sMessage, oFSO, i, sCharCode, sASCII
Dim j
sMessage = InputBox("Enter Message to Decode:")
If sMessage = "" Then WScript.Quit
set oFSO = CreateObject("Scripting.FileSystemObject")
j = 1
For i = 0 To Len(sMessage) - 1
	sCharCode = Asc(Mid(sMessage, j, 1))
	sCharCode = sCharCode - i
	sASCII = sASCII & Chr(sCharCode)
	j = j + 1
Next 

WScript.Echo sASCII


