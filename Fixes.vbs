'-------------------------------------------------------------------------
'
'Example code using WMI to fetch a list of all installed hotfixes
'
'Shawn Stugart
'=========================================================================

On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_QuickFixEngineering",,48)
For Each objItem in colItems
    WScript.Echo.vbCrLf
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CSName: " & objItem.CSName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "FixComments: " & objItem.FixComments
    Wscript.Echo "HotFixID: " & objItem.HotFixID
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "InstalledBy: " & objItem.InstalledBy
    Wscript.Echo "InstalledOn: " & objItem.InstalledOn
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "ServicePackInEffect: " & objItem.ServicePackInEffect
    Wscript.Echo "Status: " & objItem.Status
    WScript.Echo vbCrLf
Next
