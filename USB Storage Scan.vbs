strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colDevices = objWMIService.ExecQuery _
    ("Select * From Win32_USBControllerDevice")

For Each objDevice in colDevices
    strDeviceName = objDevice.Dependent
    strQuotes = Chr(34)
    strDeviceName = Replace(strDeviceName, strQuotes, "")
    arrDeviceNames = Split(strDeviceName, "=")
    strDeviceName = arrDeviceNames(1)
    Set colUSBDevices = objWMIService.ExecQuery _
        ("Select * From Win32_PnPEntity Where DeviceID = '" & strDeviceName & "'")
    For Each objUSBDevice in colUSBDevices
    	sDescription = lcase(objUSBDevice.Description)
		If sDescription = "disk drive" Then
            Wscript.Echo "USB DISK DRIVE FOUND!!!"
		End If
    Next    
Next
