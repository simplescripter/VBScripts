Option Explicit

Dim oOU, oUserCopy, oUserTemplate, arrSVAttributes
Dim arrMVAttributes, arrGroups, group, sAttrib
Dim arrValue, sResult, sValue, oNewGroup, sUserDN
Dim sUser

Const ADS_PROPERTY_UPDATE = 2
Const ADS_PROPERTY_APPEND = 3
Const E_ADS_PROPERTY_NOT_FOUND = &h8000500D

Set oOU = GetObject("LDAP://OU=Test,dc=nwtraders,dc=msft")
Set oUserCopy = oOU.Create("user", "cn=BarrAdam")
oUserCopy.Put "sAMAccountName", "barradam"
oUserCopy.SetInfo
sUserDN = oUserCopy.Get("distinguishedName")
sResult = "User " & oUserCopy.sAMAccountName & " created." & VbCrLf & VbCrLf

On Error Resume Next

Set oUserTemplate = GetObject("LDAP://cn=_Sales Template,ou=Test,dc=nwtraders,dc=msft")
arrSVAttributes = Array("description", "department", "company")
arrMVAttributes = Array("otherTelephone")
arrGroups = oUserTemplate.GetEx("memberOf")
If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
    sResult = sResult & "The memberOf attribute is not set." & VbCrLf
    Err.Clear
Else
    For Each Group in arrGroups
        sResult = sResult & AddToGroup(group, sUserDN)
    Next
End If

For Each sAttrib in arrSVAttributes
	sValue = oUserTemplate.Get(sAttrib)
	If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
		sResult = sResult & sAttrib & " property not set in template." & VbCrLf
		Err.Clear
	Else
		oUserCopy.Put sAttrib, sValue
		sResult = sResult & sAttrib & " set to " & sValue & VbCrLf
	End If
Next


If IsEmpty(arrMVAttributes) Then
 	arrMVAttributes = ""
Else
	For Each sAttrib in arrMVAttributes
		arrValue = oUserTemplate.GetEx(sAttrib)
		If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
			sResult = sResult & sAttrib & " property not set in template." & VbCrLf
			Err.Clear
		Else
			oUserCopy.PutEx ADS_PROPERTY_UPDATE, sAttrib, arrValue
			sResult = sResult & sAttrib & " updated."
		End If
	Next
End If
oUserCopy.SetInfo

WScript.Echo sResult

Function AddToGroup(sGroupName, sUser)
	Dim oNewGroup
	
	Set oNewGroup = GetObject("LDAP://" & sGroupName)
	oNewGroup.PutEx ADS_PROPERTY_APPEND, _
    	"member", Array(sUser)
	oNewGroup.SetInfo
	If Err.Number = 0 Then
		AddToGroup = "Added to " & sGroupName & "group." & vbCrLf
	Else
		AddToGroup = "Error = " & Err.Description & VbCrLf
		Err.Clear
	End If	
End Function
