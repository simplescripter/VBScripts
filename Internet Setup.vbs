Option Explicit
dim strComputer, strDNSServer
dim objWMIService, colAdapters, objAdapter

strComputer = "."
strDNSServer = Array("10.10.25.85")
Set objWMIService = GetObject("winmgmts:" & "!\\" & strComputer & "\root\cimv2")
Set colAdapters = objWMIService.ExecQuery("Select * from " _
    & "Win32_NetworkAdapterConfiguration Where IPEnabled = True")
For Each objAdapter in colAdapters
	objAdapter.SetDNSServerSearchOrder strDNSServer
	objAdapter.ReleaseDHCPLease
	objAdapter.RenewDHCPLease
Next
WScript.Echo "DONE"

