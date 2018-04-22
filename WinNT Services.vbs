strComputer = InputBox("Enter Computer Name")
set oProv = GetObject("WinNT://" & strComputer)
For Each service in oProv
    If service.Class = "Service" Then
        strResult = strResult & service.Name & vbCrLf
    End If
Next

WScript.Echo strResult