'-------------------------------------------------------------------------
'
'  Example script using WMI to fetch the name of the currently logged-on
'  user.  Does not include any error checking.
'
'  Shawn Stugart
'=========================================================================

On Error Resume Next

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFile = oFSO.OpenTextFile("C:\clients.txt", 1)

Do Until oFile.AtEndOfStream
    sCompName = oFile.Readline
    Set objWMIService = GetObject("winmgmts:\\" & sCompName & "\root\cimv2")
    If Err.Number <> 0 Then
        sResult = sResult & objitem.Name & " = " & "UNREACHABLE"
    Else
        Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem",,48)
        For Each objitem in colItems
            sResult = sResult & objitem.Name & " = " & objItem.UserName & vbCrlf
        Next
    End If
Loop

WScript.Echo sResult