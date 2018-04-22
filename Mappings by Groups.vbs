'  Example script that maps drives and printers based on Active Directory
'  group membership.
'  Note: VBScript is case sensitive when doing a string comparison.  I've chosen
'  to convert the group membership to an uppercase string.  That's why the 
'  group relative distinguished names are in all caps.

Option Explicit
Dim oADSystemInfo, oUser, oNetwork, sGroups
Const groupHR = "CN=HR" 
Const groupMan = "CN=MANAGERS" 
Const groupIT = "CN=IT" 

Set oADSystemInfo = CreateObject("ADSystemInfo") 
Set oUser = GetObject("LDAP://" & oADSystemInfo.UserName) 

' Test to see if this user is a member of more than one group.
' Note: Domain Users doesn't get counted.

If IsArray(oUser.MemberOf) Then
    sGroups = UCase(Join(oUser.MemberOf))
Else
    sGroups = UCase(oUser.MemberOf)
End If

'Map a common drive for everyone 
Set oNetwork = CreateObject("WScript.Network") 
oNetwork.MapNetworkDrive "P:", "\\denver\all" 

'Enum group memberships and map drives and printers using WSH

If InStr(sGroups, groupHR) Then
    oNetwork.MapNetworkDrive "M:", "\\denver\hr" 
	'oNetwork.AddWindowsPrinterConnection "\\<server>\<share>" 
	'oNetwork.SetDefaultPrinter "\\<server>\<share>" 
End If 

If InStr(sGroups, groupMan) Then
    oNetwork.MapNetworkDrive "I:", "\\denver\man" 
End If 

If InStr(sGroups, groupIT) Then 
    oNetwork.MapNetworkDrive "K:", "\\denver\it" 
	'oNetwork.AddWindowsPrinterConnection "\\<server>\<share>" 
	'oNetwork.SetDefaultPrinter "\\<server>\<share>" 
End If 
