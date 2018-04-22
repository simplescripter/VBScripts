'-----------------------------------------------------------------
'
' Example code demonstrating the process of sending keystrokes
' to an application.
'
' Shawn Stugart
'=================================================================

set WshShell = CreateObject("WScript.Shell")

WshShell.Run "iexplore.exe"
WScript.Sleep 1000

While WshShell.AppActivate("Microsoft Internet Explorer") = FALSE
	wscript.sleep 1000
Wend

WshShell.AppActivate "Microsoft Internet Explorer"
WScript.Sleep 2000

WshShell.SendKeys "%F"
WScript.Sleep 200

WshShell.SendKeys "o"
WScript.Sleep 200

WshShell.SendKeys "www.yahoo.com"

WshShell.Sendkeys "{ENTER}"
Wscript.Sleep 1500

WshShell.SendKeys "%F"
WScript.Sleep 200

WshShell.SendKeys "o"
WScript.Sleep 200

WshShell.SendKeys "www.espn.com"

WshShell.Sendkeys "{ENTER}"
Wscript.Sleep 5500

WshShell.SendKeys "%F"
WScript.Sleep 200

WshShell.SendKeys "o"
WScript.Sleep 200

WshShell.SendKeys "www.nhcolorado.com"

WshShell.Sendkeys "{ENTER}"
Wscript.Sleep 1500
