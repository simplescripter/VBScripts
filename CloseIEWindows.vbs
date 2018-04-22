Option Explicit

Dim oShell, oWindows, sURL, iCount
Dim i, oIE, window

Set oShell = CreateObject("Shell.Application")
Set oWindows = oShell.Windows

sURL = "http://www.nhcolorado.com/denver/"
iCount = 0

On Error Resume Next

For i = 0 To oWindows.count - 1
	If InStr(1, oWindows.item(i).LocationURL, sURL) Then
		iCount = iCount + 1
    End If
Next
If iCount > 2 Then
	WScript.Echo iCount
  	WScript.Echo "Too many windows!! Close 'em!"
  	For Each window In oWindows
  		If window.locationUrl = sURL then
  			window.Quit()
   		End If
  	Next
End If