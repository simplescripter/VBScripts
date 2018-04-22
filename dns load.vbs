Set oShell = CreateObject("WScript.Shell")
k = 1
For i = 0 To 255
	For j = 1 To 254
		sIP = "192.168." & i & "." & j
		oShell.Exec "dnscmd den-dc1.contoso.msft /recordadd contoso.msft comp" & k & " A " & sIP
		k = k + 1
		WScript.StdOut.Write "."
	Next
Next