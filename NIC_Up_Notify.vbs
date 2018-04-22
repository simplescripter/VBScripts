sSubnet = "192.168.105"
sSuccessMessage = "You're ready to roll!"

Set oWMI = GetObject("WinMgmts:\\")
Do
	Set colIPs = oWMI.ExecQuery _
    	("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
	For Each sIP in colIPs
	    If Not IsNull(sIP.IPAddress) Then 
	        For i=LBound(sIP.IPAddress) to UBound(sIP.IPAddress)
	            If Instr(sIP.IPAddress(i),sSubnet) Then
	            	WScript.Echo sSuccessMessage
	            	Exit Do
	            End If
			Next
	    End If
	Next
	WScript.Sleep 1000
Loop
WScript.Quit