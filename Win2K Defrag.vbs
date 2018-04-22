
'This script launches defrag and sends keys to the UI in order to automate the defrag
'process.

set WshShell = CreateObject("WScript.Shell")

'Launch Defrag from the command line and wait for a second
WshShell.Run "dfrg.msc"
WScript.Sleep 1000

'Wait until the application has loaded - Check every second
While WshShell.AppActivate("Disk Defragmenter") = FALSE
	wscript.sleep 1000
Wend

'Bring the application to the foreground
WshShell.AppActivate "Disk Defragmenter"
WScript.Sleep 200

'Send an ALT-A key to bring down the degrag menu
WshShell.SendKeys "%A"
WScript.Sleep 200

'Send a D to start the defrag
WshShell.SendKeys "D"

'Wait until the defrag is completed - Check for window every 5 seconds
While WshShell.AppActivate("Defragmentation Complete") = FALSE
	wscript.sleep 5000
Wend

'Bring the msgbox to the foreground
WshShell.AppActivate "Defragmentation Complete"
WScript.Sleep 200

'Send a tab key to move the focus from View Report button to the Close Button
WshShell.Sendkeys "{TAB}"
Wscript.Sleep 500

'Send key to Close the Defragmentation Complete window
WshShell.Sendkeys "{ENTER}"
Wscript.Sleep 500

'Send and ALT-F4 to Close the Defrag program
WshShell.Sendkeys "%{F4}"
