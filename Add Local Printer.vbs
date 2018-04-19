strComputer = "."

Set objWMIService = GetObject("winmgmts:")
Set objNewPort = objWMIService.Get _
    ("Win32_TCPIPPrinterPort").SpawnInstance_
objNewPort.Name = "IP_10.10.25.25"
objNewPort.Protocol = 1
objNewPort.HostAddress = "10.10.25.25"
objNewPort.PortNumber = "9999"
objNewPort.SNMPEnabled = False
objNewPort.Put_


Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set objPrinter = objWMIService.Get("Win32_Printer").SpawnInstance_
objPrinter.DriverName = "HP LaserJet 4000 Series PS"
objPrinter.Caption = "Scripted Printer"
objPrinter.Default = True
objPrinter.Direct = False
objPrinter.Local = True
objPrinter.PortName   = "IP_10.10.25.25"
objPrinter.DeviceID   = "Terminal Server User Printer"
objPrinter.Shared = False
objPrinter.Location = "USA/Colorado/DTC/"
objPrinter.Network = False
objPrinter.Put_