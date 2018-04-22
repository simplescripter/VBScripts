Dim WriteConsonant	
Dim nbRnd		
Dim i, tmp

'Register AutoIt and temporarily freeze mouse and keyboard
Set oShell = WScript.CreateObject("WScript.Shell")
oShell.Run "regsvr32.exe c:\autoitx.dll"
WScript.Sleep 2000
Set oAutoIt = WScript.CreateObject("AutoItX.Control")
WScript.Sleep 2000
oAutoIt.BlockInput "on"
oAutoIt.WinClose "RegSvr32", ""
WScript.Sleep 2000

 If oAutoIt.IfWinExist ("Windows", "A duplicate name") = 1 Then
     oAutoIt.WinClose "Windows", "A duplicate name"
 End If

oShell.Popup "A SCRIPT IS RUNNING, PLEASE DO NOT USE THE MOUSE OR KEYBOARD UNTIL ALL SCRIPTS HAVE COMPLETED",10,"ATTENTION!",16

'Create a somewhat pronouncable password 6 to 11 letters
Const strVowel = "AEIOU"		   
Const strConsonant = "BCDFGHJKLMNPRSTV"	   
Const strDoubleConsonant = "CDFGLMNPRST" 
 
GeneratePassword = ""
WriteConsonant = False
i = 0
upper = 11
lower = 6

Randomize
LenPassword = Int((upper - lower + 1) * Rnd + lower)

Randomize
For i = 0 To LenPassword	
nbRnd = Rnd
    	
    	If GeneratePassword <> "" And (WriteConsonant = False) And (nbRnd < 0.10) Then
    	     tmp = Mid(strDoubleConsonant, Int(Len(strDoubleConsonant) * Rnd + 1), 1)
    	     tmp = tmp & tmp
    	     i = i + 1
    	     WriteConsonant = True
    	ElseIf (WriteConsonant = False) And (nbRnd < 0.90) Then
    	     tmp= Mid(strConsonant, Int(Len(strConsonant) * Rnd + 1), 1)
    	     WriteConsonant = True
     	Else tmp = Mid(strVowel,Int(Len(strVowel) * Rnd + 1), 1)
    	     WriteConsonant = False
    	End If
        GeneratePassword = GeneratePassword & tmp
Next    	
If Len(GeneratePassword) > LenPassword Then
     GeneratePassword = Left(GeneratePassword, LenPassword)
End If 

oAutoIT.ClipPut(GeneratePassword)
Dim proctype
Set WshSysEnv = oShell.Environment("SYSTEM")
proctype = Right (WshSysEnv("PROCESSOR_IDENTIFIER"), 12)

'Pause for plug and play to run
Do While count < 30
  	If oAutoIt.IfWinExist ("System Settings Change", "Windows 2000 has finished") = 1 Then
		oShell.AppActivate "System"
		WScript.Sleep 1000
		oAutoIt.Send "!n"
		Call Finish

	ElseIf oAutoIt.IfWinExist ("Found New", "This wizard helps you") = 1 Then
		Call Finish

	Else  
		oAutoIt.Sleep 10000
		count = count + 1
  	End If 	
Loop

Call Finish

'Rename, install drivers and reboot
Sub Finish()
WScript.Sleep 2000
oAutoIt.BlockInput "off"

 If oAutoIt.IfWinExist ("Windows", "A duplicate name") = 1 Then
     oAutoIt.WinClose "Windows", "A duplicate name"
 End If

 If proctype = ("AuthenticAMD") THEN
		Call AMD
 End If	

WScript.Sleep 2000
oShell.Run "control"
WScript.Sleep 4000
oShell.AppActivate "Control"
WScript.Sleep 2000
oShell.SendKeys "{s 5} {ENTER}"
WScript.Sleep 3000
oAutoIt.WinActivate "System", ""
WScript.Sleep 2000
oShell.SendKeys "{TAB 3}#{RIGHT}"
WScript.Sleep 3000
oShell.SendKeys "%r"
WScript.Sleep 3000
oShell.SendKeys "^v"
WScript.Sleep 3000
oShell.SendKeys "{ENTER}"
WScript.Sleep 3000
oShell.SendKeys "{ENTER}"
WScript.Sleep 3000
oShell.AppActivate "Control"
WScript.Sleep 2000
oShell.SendKeys "%f c"
WScript.Sleep 3000
oShell.SendKeys "{ENTER}"
WScript.Sleep 2000
oshell.SendKeys "%n"

'Detect video card and install drivers if S3Trio3D2X
Dim vidtype
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_VideoController",,48)

For Each objItem in colItems
   
        vidtype = Left (objItem.Description, 17)
	
Next
  
  If vidtype = ("S3 Inc. Trio3D/2X") Then
        	oShell.Run "C:\3D2X.exe"
	oAutoIt.SetKeyDelay 700
	oAutoIt.Sleep 3000
	oAutoIt.WinWaitActive "3D2X", ""
	oAutoIt.Sleep 1000
	oAutoIt.Send "{ENTER}"
	oAutoIt.Sleep 3000
	oAutoIt.Send "{ENTER}"
  End If

oAutoIt.Shutdown"3"
End Sub

	Sub AMD()
		oAutoIt.SetKeyDelay 500
		oAutoIt.Sleep 2000
		oShell.Run "C:\SndVid.exe"
		oAutoIt.Sleep 3000
		oAutoIt.WinActivate "ZipCentral", ""
		oAutoIt.Sleep 2000
		oAutoIt.Send "!e#c:\sndvid#{TAB 5}#{ENTER}"
		oAutoIt.WinWaitActive "Finished", ""
		oAutoIt.Sleep 2000
		oAutoIt.Send "{ENTER}"
		oAutoIt.WinActivate "Found", ""
		oAutoIt.Sleep 1000
		oAutoIt.Send "!n"
		oAutoIt.Sleep 2000
		oAutoIt.Send "!n"
		oAutoIt.Sleep 2000
		oAutoIt.Send "!n"
		oAutoIt.Sleep 2000
		oAutoIt.Send "C:\SndVid#{ENTER}"
		oAutoIt.WinWaitActive "Found", "Windows found a driver"
		oAutoIt.Sleep 2000
		oAutoIt.Send "!n"
		oAutoIt.Sleep 2000
		oAutoIt.Send "!y"
		oAutoIt.Sleep 3000
		oAutoIt.WinWaitActive "Found", "Windows has finished"
		oAutoIt.Sleep 1000
		oAutoIt.Send "{ENTER}"
		oAutoIt.Sleep 3000
		oAutoIt.Send "!n"
		oAutoIt.Sleep 2000
		oAutoIt.Send"!n"
		oAutoIt.Sleep 2000
		oAutoIt.Send "!n"
		oAutoIt.Sleep 2000
		oAutoIt.Send "{ENTER}"
		oAutoIt.Sleep 2000
		oAutoIt.Send "!n"
		oAutoIt.Sleep 4000
		oAutoIt.Send "{ENTER}"
		oAutoIt.Sleep 3000
		oAutoIt.Send "!n"
		oAutoIt.Sleep 2000
		oShell.Run "C:\Res.exe"
		oAutoIt.Sleep 2000
		oAutoIt.WinWaitActive "Res", ""
		oAutoIt.Sleep 1000
		oAutoIt.Send "{ENTER}"
		oAutoIt.Sleep 2000
	End Sub

WScript.Quit