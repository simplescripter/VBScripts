' *****************************************************************************
'
' This script was adapted from the Scripting Guys at Microsoft.
'
' A note on running this script on XP SP2/Win2K3 SP1 and later systems:
'
'	Any script using WMI over the network with the above-mentioned OS and 
'	service pack (and later) is likely to encounter issues, both from the
'	Windows Firewall and the enhanced DCOM security settings now in place.
'	This script uses an asynchronous call to monitor processes on remote
'	systems, adding to the configuration hassle.  In particular, you'll have
'	problems running this script in a workgroup environment.  Here are the
'	steps I took to get this script working:
'	
'		1. Enable Remote Administration through the Windows Firewall on the
'			systems being monitored.  This can be done through Local Policy
'			(or Group Policy in a domain), or with a direct Registry edit.
'			The WMI connection to the remote client uses DCOM, which relies on
'			RPCs, and allowing Remote Administration through the Windows
'			Firewall will take care of this requirement.
'		2. This script uses a SINK object with an Asynchronous call to the
'			systems being monitored.  In an asynchronous call, the management
'			system (where the script is running) makes an initial DCOM/RPC
'			connection to the target (the system being monitored), then the
'			target makes an ANONYMOUS DCOM/RPC call back to the management
'			system.  For this to work, you need to make some changes ON THE
'			MACHINE WHERE THE SCRIPT IS RUNNING:
'
'			A. Open the DCOM port.  This is TCP port 135
'			B. Create a program exception in the firewall for C:\Windows\
'				System32\wbem\unsecapp.exe.
'			C. Use dcomcnfg to change the default DCOM permissions and allow
'				the ANONYMOUS LOGON account the following permissions:
'				1. On the COM Security tab, under Access Permissions, Select
'					Edit Limits... and grant ANONYMOUS LOGON the Remote
'					Access permission.
'				2. REBOOT!!!
'
' Shawn W Stugart
' 12-1-2005
'
'******************************************************************************

On Error Resume Next

Dim arrComputers()

CheckCScript

g_arrTargetProcs = Array("calc.exe")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oInputFile = oFSO.OpenTextFile("C:\Script\Clients.txt",1)
Set SINK = WScript.CreateObject("WbemScripting.SWbemSink","SINK_")
intSize = 0
Do Until oInputFile.AtEndOfStream
	ReDim Preserve arrComputers(intSize)
	arrComputers(intSize) = oInputFile.ReadLine
	strComputer = arrComputers(intSize)
  	WScript.Echo vbCrLf & "Host: " & strComputer
  	intKP = KillProcesses(strComputer)
  	If intKP = 0 Then
    	TrapProcesses strComputer
 	 Else
    	WScript.Echo vbCrLf & "  Unable to monitor target processes."
  	End If
  	intSize = intSize + 1
Loop

Wscript.Echo VbCrLf & _
 "     -----------------------------------------------------------------" & _
 VbCrLf & VbCrLf & "In monitoring mode ..."
 
Do
   WScript.Sleep 1000
Loop

'******************************************************************************

Function KillProcesses(strHost)
'Terminate specified processes on specified machine.

On Error Resume Next

strQuery = "SELECT * FROM Win32_Process"
intTPFound = 0
intTPKilled = 0

Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" & strHost & "\root\cimv2")
If Err = 0 Then
  WScript.Echo vbCrLf & "  Searching for target processes."
  Set colProcesses = objWMIService.ExecQuery(strQuery)
  For Each objProcess in colProcesses
    For Each strTargetProc In g_arrTargetProcs
      If LCase(objProcess.Name) = LCase(strTargetProc) Then
        intTPFound = intTPFound + 1
        WScript.Echo "  " & objProcess.Name
        intReturn = objProcess.Terminate
        If intReturn = 0 Then
          WScript.Echo "    Terminated"
          intTPKilled = intTPKilled + 1
        Else
          WScript.Echo "    Unable to terminate"
        End If
      End If
    Next
  Next

  WScript.Echo "  Target processes found: " & intTPFound
  If intTPFound <> 0 Then
    WScript.Echo "  Target processes terminated: " & intTPKilled
  End If
  intTPUndead = intTPFound - intTPKilled
  If intDiff <> 0 Then
    WScript.Echo "  ALERT: Target processes not terminated: " & intTPUndead
  End If
  KillProcesses = 0
Else
  HandleError(strHost)
  KillProcesses = 1
End If

End Function

'******************************************************************************

Sub TrapProcesses(strHost)

On Error Resume Next

strAsyncQuery = "SELECT * FROM __InstanceCreationEvent WITHIN 1 " & _
 "WHERE TargetInstance ISA 'Win32_Process'"

'Connect to WMI.
Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" & strHost & "\root\cimv2")
If Err = 0 Then
'Trap asynchronous events.
  objWMIService.ExecNotificationQueryAsync SINK, strAsyncQuery
  If Err = 0 Then
    WScript.Echo vbCrLf & "  Monitoring target processes."
  Else
    HandleError(strHost)
    WScript.Echo "  Unable to monitor target processes."
  End If
Else
  HandleError(strHost)
  WScript.Echo "  Unable to monitor target processes."
End If

End Sub

'******************************************************************************

Sub SINK_OnObjectReady(objLatestEvent, objAsyncContext)
'Trap asynchronous events.

For Each strTargetProc In g_arrTargetProcs
  If LCase(objLatestEvent.TargetInstance.Name) = LCase(strTargetProc) Then
    Wscript.Echo VbCrLf & "Target process on: " & _
     objLatestEvent.TargetInstance.CSName
    Wscript.Echo "  Name: " & objLatestEvent.TargetInstance.Name
    Wscript.Echo "  Time: " & Now
'Terminate process.
    intReturn = objLatestEvent.TargetInstance.Terminate
    If intReturn = 0 Then
      Wscript.Echo "  Terminated process."
    Else
      Wscript.Echo "  Unable to terminate process. Return code: " & intReturn
    End If
  End If
Next

End Sub

'******************************************************************************

Sub HandleError(strHost)
'Handle errors.

strError = VbCrLf & "  ERROR on " & strHost & VbCrLf & _
 "  Number: " & Err.Number & VbCrLf & _
 "  Description: " & Err.Description & VbCrLf & _
 "  Source: " & Err.Source
WScript.Echo strError
Err.Clear

End Sub

'******************************************************************************
Sub CheckCScript()
	hostname = lcase(right(WSCript.Fullname, 11))
	If hostname = "wscript.exe" Then
	    WScript.Echo "This script requires cscript.exe"
	    iHelp = MsgBox("At a command line, run " & Chr(34) _
	    	& "cscript " & WScript.ScriptFullName & Chr(34) _
	    	& vbCrLf & vbCrLf & "Do you want help?", vbYesNo)
	    If iHelp = vbYes Then
	    	Set oShell = CreateObject("WScript.Shell")
	    	oShell.Run("cmd.exe")
	    	WScript.Sleep 500
	    	oShell.SendKeys "cscript.exe " & Chr(34) & WScript.ScriptFullName & Chr(34)
	    	WScript.Quit
	    Else
	    	WScript.Quit
	    End If
	End If
End Sub

