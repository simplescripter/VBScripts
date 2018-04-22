'******************************************************************************
'FwRemoteAdminEnable.vbs
'Author: Peter Costantini, The Microsoft Scripting Guys
'Date: 9/1/04
'Version: 1.0
'This script enables remote administration for Windows Firewall.
'Remote administration is disabled by default.
'******************************************************************************

On Error Resume Next
'Create the firewall manager object.
Set objFwMgr = CreateObject("HNetCfg.FwMgr")
If Err <> 0 Then
  WScript.Echo "Unable to connect to Windows Firewall."
  WScript.Quit
End If
'Get the current profile for the local firewall policy.
Set objProfile = objFwMgr.LocalPolicy.CurrentProfile
WScript.Echo VbCrLf & "Windows Firewall"

'Get remote admin settings.
Set objRemoteAdminSettings = objProfile.RemoteAdminSettings
WScript.Echo VbCrLf & "Current Remote Administration Settings:"
WScript.Echo "Enabled: " & objRemoteAdminSettings.Enabled
If objRemoteAdminSettings.Scope = 0 Then
  strScope = "All" 'Default
ElseIf objRemoteAdminSettings.Scope = 1 Then
  strScope = "Local Subnet"
Else
  strScope = "UNKNOWN"
End If
WScript.Echo "Scope: " & strScope
'If remote administration not enabled, enable it.
If objRemoteAdminSettings.Enabled = False Then
  objRemoteAdminSettings.Enabled = True
  WScript.Echo VbCrLf & "Remote Administration enabled."
End If

