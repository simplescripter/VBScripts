sComputer = "."
Set oWMI = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate,(Shutdown)}!\\" & sComputer & "\root\cimv2")
Set colComputers = oWMI.ExecQuery _
    ("Select * from Win32_ComputerSystem")
For Each oComputer in colComputers
    err = oComputer.Rename("Instructor")
Next

Set colOperatingSystems = oWMI.ExecQuery _
    ("Select * from Win32_OperatingSystem")
For Each oOperatingSystem in colOperatingSystems
    oOperatingSystem.Reboot()
Next
	