set oWMIShares = GetObject("WinMgmts:").InstancesOf("Win32_Share")
set oShell = CreateObject("WScript.Shell")
For Each oShare in oWMIShares
    Return = oShell.Run("%comspec% /c net share " & oShare.Name & " /cache:no", 0)
Next
WScript.Echo "Finished."