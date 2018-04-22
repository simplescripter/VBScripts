'#----- #
'# Author: Mr. Lee #
'# Use code at will... #
'#----- #

' --Begin Instructions----
' This is just a section of my Logon Script.
' Create a csv file. 
' Name it Grouplist.csv, or if you name it something else, change code below.
' Make the names of the group the Prewindows2000 style. IE. "Domain Users"
' Group Name, Drive Letter, UNC Path to Share Like this: 
' Domain Users,J:,\\TokyoServer\Common
' Domain Admins,K:,\\Tokyoserver\Adminshare$
' Finance,F:,\\TokyoFNSRV\Finance

'----Begin Script--- 

Dim GroupList
Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell") 
Set WshNetwork = WScript.CreateObject("WScript.Network") 

GetGroupInfo()

LogonPath = fso.GetParentFolderName(WScript.ScriptFullName)
'**************************************Group Mappings Based on Grouplist.csv*********************************
If fso.FileExists(logonpath&"\Grouplist.csv") Then
	Set grplist = Fso.OpenTextFile(logonpath&"\Grouplist.csv")
	' make File into an Array
	aGroup = Split(grplist.Readall,vbcrlf)
	For I = 0 to UBound(GroupList) ' Check Every Group Membership the user is in (populated into Grouplist)
		grpname = Grouplist(i)
		For x = 0 to UBound(aGroup) ' Read the entire CSV to make sure all drives are mapped for each Group
			mapline = agroup(x)
			If InStr(LCase(mapline),LCase(grpname)) Then ' If you're in the group
				mapline = Mid(mapline,InStr(mapline,",")+1) ' Remove the GroupName from the line
				Drive = Left(mapline,InStr(mapline,",")-1) ' Extract Drive Letter
				Path = Mid(mapline,InStr(mapline,",")+1) ' Extract the path
				If fso.DriveExists(drive) <> True Then ' If The Drive is not already mapped
					WshNetwork.MapNetworkDrive drive,path ' Map The Drive
					wscript.sleep 1000
				End If
			End If
		Next
	Next
End If


Sub GetGroupInfo
	'msgbox("IN GroupInfo")
	Set UserObj = GetObject("WinNT://" & wshNetwork.UserDomain & "/" & WshNetwork.UserName)
	Set Groups = UserObj.groups

	For Each Group In Groups
		GroupCount = GroupCount + 1
	Next

	ReDim GroupList(GroupCount -1)
	i = 0
	For Each Group In Groups
		GroupList(i) = Group.Name
		i = i + 1
		'msgbox(Group.Name)
	Next
End Sub 
