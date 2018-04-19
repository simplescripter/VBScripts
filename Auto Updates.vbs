'------------------------------------------------------------------------
'  This script allows an administrator to configure the necessary
'  registry entries on a local or remote client to support Automatic
'  Updates with SUS in a non-Active Directory environment.
'
'  Edit the variables appropriately for your deployment.  
'
'  Shawn Stugart
'  10-29-2004
'========================================================================
Option Explicit
const HKEY_LOCAL_MACHINE = &H80000002

Dim dwNoAutoUpdate, dwAUOptions, dwScheduledInstallDay, dwScheduledInstallTime
Dim dwUseWUServer, dwRescheduleWaitTime, dwNoAutoRebootWithLoggedOnUsers
Dim szWUServer, szWUStatusServer, oShell, strComputer, oScriptExec, oReg
Dim strPingStdOut, strKeyPath, strKeyPath1, strKeyPath2, arrSubkeys, strKeys

On Error Resume Next

dwNoAutoUpdate = 0 'Set to 1 to disable Automatic Updates
dwAUOptions = 4 ' A value of 4 sets the client to auto download and auto install updates
dwScheduledInstallDay = 0 ' Client installs new updates daily
dwScheduledInstallTime = 3 ' Time (3:00 am) client installs updates
dwUseWUServer = 1 ' Use Windows Update or an SUS server
dwRescheduleWaitTime = 5 ' Reschedule missed update installations for 5 minutes after Auto Updates starts
dwNoAutoRebootWithLoggedOnUsers = 1 ' Don't automatically reboot when users are logged on
szWUServer = "http://server" ' URL of the SUS server
szWUStatusServer = "http://server" ' URL of the SUS statistics server (usually the SUS server itself)

set oShell = CreateObject("WScript.Shell")
strComputer = InputBox("Enter the name of the machine on which to " _
    & "enable and configure Automatic Updates","ENTER MACHINE NAME")
If strComputer = "" Then WScript.Quit
Set oScriptExec = oShell.Exec("ping -n 1 -w 1000 " & strComputer)
strPingStdOut = Lcase(oScriptExec.StdOut.ReadAll)
If InStr(strPingStdOut, "reply from ") Then
    Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _
        & strComputer & "\root\default:StdRegProv")
    strKeyPath =  "SOFTWARE\Policies\Microsoft\Windows"
	strKeyPath1 = "SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
	strKeyPath2 = "SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
	oReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys
	strKeys = Join(arrSubkeys)
	If Not(InStr(strKeys,"AU")) Then
		oReg.CreateKey HKEY_LOCAL_MACHINE,strKeyPath2 & "\AU"
	End If
	oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath1,"NoAutoUpdate",dwNoAutoUpdate
	oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath1,"AUOptions",dwAUOptions
	oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath1,"ScheduledInstallDay",dwScheduledInstallDay
	oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath1,"ScheduledInstallTime",dwScheduledInstallTime
	oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath1,"UseWUServer",dwUseWUServer
	oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath1,"RescheduleWaitTime",dwRescheduleWaitTime
	oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath1,"NoAutoRebootWithLoggedOnUsers",dwNoAutoRebootWithLoggedOnUsers
	oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath2,"WUServer", szWUServer
	oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath2,"WUStatusServer", szWUStatusServer
	If Err.Number = 0 Then
	    WScript.Echo "Automatic Updates has been configured " _
	        & "on " & strComputer
	Else
	    WScript.Echo "Configuration of Automatic Updates FAILED on " & strComputer
	End If
End If
