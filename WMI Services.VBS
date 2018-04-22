strComputer = InputBox("Enter Computer Name:")
For Each service in GetObject("WinMgmts:\\" & strComputer).InstancesOf("Win32_Service")
    strResult = strResult & service.Name & vbCrLf
Next

WScript.Echo strResult 