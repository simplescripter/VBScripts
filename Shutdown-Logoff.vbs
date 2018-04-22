Dim oFSO, oTS, sClient, oWindows, oLocator, oConnection, oSys
Dim sUser, sPassword

'set remote credentials
sUser = "Administrator"
sPassword = "password"

'open list of client names
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oTS = oFSO.OpenTextFile("C:\clients.txt")

Do Until oTS.AtEndOfStream
 
 'get next client name 
 sClient = oTS.ReadLine
 
 'get WMI locator
 Set oLocator = CreateObject("WbemScripting.SWbemLocator")

 'Connect to remote WMI
 Set oConnection = oLocator.ConnectServer(sClient, _
   "root\cimv2", sUser, sPassword)

  'issue shutdown to OS
 ' 4 = force logoff
 ' 5 = force shutdown
 ' 6 = force rebooot 
 ' 12 = force power off
 Set oWindows = oConnection.ExecQuery("Select " & _
   "Name From Win32_OperatingSystem")
 For Each oSys In oWindows
   oSys.Win32ShutDown(5)
 Next 

Loop

'close the text file
oTS.Close
WScript.Echo "All done!"

