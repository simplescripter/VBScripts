Const HKEY_CURRENT_USER = &H80000001 
'On Error Resume Next

strComputer = "." 
Set objRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv") 
  
strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections" 
strValueName = "DefaultConnectionSettings" 
sResult = ""
rtn1 = objRegistry.GetBinaryValue(HKEY_CURRENT_USER, strKeyPath, strValueName, binValue) 
If rtn1 = 0 Then
	For i = 0 To 7
		sResult = sResult & binValue(i) & ","
	Next
	sResult = sResult & "9" & ","
	For j = 9 To uBound(binValue)
		If j <> UBound(binValue) Then
			sResult = sResult & binValue(j) & ","
		Else
			sResult = sResult & binValue(j)
		End If
	Next
Else
	WScript.Echo "An error occurred reading the value from the registry."
	WScript.Quit
End If

arrSplit = Split(sResult, ",")
For k = 0 To UBound(arrSplit)
	ReDim Preserve arrBinVal(k)
	arrBinVal(k) = arrSplit(k)
Next

rtn2 = objRegistry.SetBinaryValue(HKEY_CURRENT_USER, strKeyPath, strValueName, arrBinVal)
If rtn2 <> 0 Then
	WScript.Echo "An error occurred writing the value to the registry."
	WScript.Quit
End If