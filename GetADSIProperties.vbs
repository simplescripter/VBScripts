Option Explicit
Const ADS_SECURE_AUTHENTICATION = &H1
Dim oMyDomain, oRootDSE, oMyDS, iCount
Dim iLoop, sPropList, sProp, sPropName, hostname

CheckCScript

Set oRootDSE = GetObject("LDAP://RootDSE") 
Set oMyDS = Getobject("LDAP:")
Set oMyDomain = oMyDS.OpenDSObject("LDAP://" _
    & oRootDSE.Get("defaultNamingContext"), _
    "administrator", "password", ADS_SECURE_AUTHENTICATION) 
oMyDomain.GetInfo
iCount = oMyDomain.PropertyCount
sPropList = "There are " & iCount _
    & " values in the local property cache for Domain: " _
    & oMyDomain.ADSPath & vbCRLF
WScript.Echo sPropList
sPropList = ""
On Error Resume Next
For iLoop = 0 to (iCount) - 1
	Set sProp = oMyDomain.Item(CInt(iLoop))
	sPropName = sProp.Name
	sPropList = sPropList & sPropName & ": " _
	    & oMyDomain.Get(sPropName) & vbCrLf
	If err.number <> 0 then
	    err.clear
	    sPropList = sPropList & sPropName & ": " _
	    & "<MultiValued>" & vbCrLf
	End If
Next

WScript.Echo sPropList

Sub CheckCScript()
	hostname = lcase(right(WSCript.Fullname, 11))
	If hostname = "wscript.exe" Then
	    WScript.Echo "This script requires cscript.exe"
	    WScript.Echo "At a command line, run " & Chr(34) _
	        & "cscript " & WScript.ScriptFullName & Chr(34)
	    WScript.Quit
	End If
End Sub