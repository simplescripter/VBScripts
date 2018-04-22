'-------------------------------------------------------------------------
'
'  Example script using WMI to fetch the name of the currently logged-on
'  user.  Does not include any error checking.
'
'  Shawn Stugart
'=========================================================================

On Error Resume Next
strComputer = "shade"
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem",,48)
For Each objitem in colItems
    Wscript.Echo "Currently logged on user at " & objitem.Name & " is " & objItem.UserName
Next