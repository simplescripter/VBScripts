Option Explicit
dim oAdminUsers, oAdminMember
dim oPUsers, oPUMember
dim sAdminResult, sPUResult, sComputer

sComputer = InputBox("Enter computer name")
Set oAdminUsers = GetObject("WinNT://" & sComputer & "/Administrators")
For each oAdminMember in oAdminUsers.Members
    sAdminResult = sAdminResult & oAdminMember.Name & vbCrLf
Next
Set oPUsers = GetObject("WinNT://" & sComputer & "/Power Users")
For each oPUMember in oPUsers.Members
    sPUResult = sPUResult & oPUMember.Name & vbCrLf
Next

WScript.Echo "Administrators on " & sComputer & ":" & vbCrLf
WScript.Echo sAdminResult & vbCrLf
WScript.Echo "Power Users on " & sComputer & ":" & vbCrLf
WScript.Echo sPUResult & vbCrLf