sComputer = "."
Set oWMI = GetObject("WinMgmts://" & sComputer)
Set oDrives = oWMI.ExecQuery("Select * From Win32_Volume Where Name = 'F:\\'")
For Each oDrive in oDrives
	oDrive.DriveLetter = "P:"
	oDrive.Put_
Next