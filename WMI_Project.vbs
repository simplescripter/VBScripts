Option Explicit

Dim oWMI, oConsumer, sPathToMOF, errReturn
Dim oWMIService, iProcessID, sComputer
Dim oNet, oFSO, oFile, myInstance, myFilter
Dim myBinding, sProcessName, iWait, sScriptName
Dim sResult

sPathToMOF = "C:\Windows\System32\wbem\scrcons.mof"
sComputer = InputBox("Computer:")
sProcessName = InputBox("Process:")
sScriptName = "NO_Calc"

On Error Resume Next
Set oWMI = GetObject("WinMgmts://" & sComputer)
If Err <> 0 Then
	WScript.Echo "Unreachable."
	WScript.Quit
End If
Set oConsumer = oWMI.Get("ActiveScriptEventConsumer")
If Err <> 0 Then
	Err.Clear
	Set oWMIService = oWMI.Get("Win32_Process")
	errReturn = oWMIService.Create("C:\Windows\System32\" _
		& "wbem\mofcomp -n:root\CIMV2 " & sPathToMOF, Null, Null, iProcessID)
	If errReturn = 0 Then
		WScript.Sleep 1000 'give mofcomp a chance to finish installing the consumer
		Do
			Set oConsumer = oWMI.Get("ActiveScriptEventConsumer")
			If Err.Number = 0 Then
				Exit Do
			Else
				LetsWait
				Err.Clear
			End If
		Loop
		WriteKillIt(sComputer)
	Else
		WScript.Echo "Failed. Code = " & errReturn
		Err.Clear
	End If
Else
	WriteKillIt(sComputer)
End If

Set myInstance = oConsumer.spawninstance_
	myInstance.name = sScriptName
	myInstance.ScriptingEngine = "VBScript"
	myInstance.ScriptFileName = "C:\Windows\System32\" & sScriptName & ".vbs"
	myInstance.ScriptText = NULL
	myInstance.KillTimeout = 30
	myInstance.put_
If Err = 0 Then
	sResult = sResult & "Success! " & sScriptName & " instance created" & vbCrLf
Else
	sResult = sResult & "Logical consumer failed: " & hex(Err.number) & ", " & err.Description & VbCrLf
End If

Set myFilter = oWMI.GET("__EventFilter").spawnInstance_
	myFilter.name = sScriptName & "Filter"
	myFilter.Query = "SELECT * FROM __InstanceCreationEvent WITHIN 5 WHERE TargetInstance isa 'Win32_Process' AND TargetInstance.name='" & sProcessName & "'"
	myFilter.QueryLanguage = "WQL"
	myFilter.put_
If Err = 0 Then
	sResult = sResult & "Success! " & sScriptName & "Filter instance created" & vbCrLf
Else
	sResult = sResult & "Filter creation failed: " & hex(err.number) & ", " & Err.Description & vbCrLf
End If

Set myBinding = oWMI.Get("__FilterToConsumerBinding").spawnInstance_
    myBinding.Consumer 		= "\\.\root\cimv2:ActiveScriptEventConsumer.name=""" & sScriptName & """"
    myBinding.Filter		= "__EventFilter.name=""" & sScriptName & "Filter"""
    myBinding.DeliverSynchronously=False
	myBinding.put_
If Err = 0 Then
	sResult = sResult & "Success! Binding instance created" & vbCrLf
Else
	sResult = sResult & "Binding failed: " & hex(err.number) & ", " & err.Description & vbCrLf
End If

WScript.Echo sResult

Function WriteKillIt(sComp)
	On Error Resume Next
	Set oNet = CreateObject("WScript.Network")
	oNet.RemoveNetworkDrive "X:", True
	If Err.Number <> 0 Then
		Err.Clear
	End If
	oNet.MapNetworkDrive "X:", "\\" & sComp & "\C$"
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	If oFSO.FileExists ("X:\Windows\System32\" & sScriptName & ".vbs") Then
		WriteKillIt = 1
		oNet.RemoveNetworkDrive "X:", True
	Else
		Set oFile = oFSO.CreateTextFile("X:\Windows\System32\" & sScriptName & ".vbs")
		oFile.WriteLine ("Set oWMI = GetObject(" & Chr(34) & "WinMgmts:" & Chr(34) & ")")
		oFile.WriteLine ("Set colProcessList = oWMI.ExecQuery (" & Chr(34) & "SELECT * FROM Win32_Process WHERE Name = " & Chr(34) & " & LCase(""'" & sProcessName & "'"")" & ")")
		oFile.WriteLine ("For Each objProcess in colProcessList")
		oFile.WriteLine ("objProcess.Terminate()")
		oFile.WriteLine ("Next")
		oFile.Close
		oNet.RemoveNetworkDrive "X:", True
	End If
End Function

Sub LetsWait()
	iWait = MsgBox("Failed binding to ActiveScriptEventConsumer.  Mofcomp may not be finished. " _
		& "Do you want to wait a few more seconds?", vbYesNo+vbQuestion, "ERROR.")
	If iWait = vbYes Then
		WScript.Sleep 3000
	Else
		WScript.Quit
	End If
End Sub