'************* CREATE SHARE AND APPLY NTFS RIGHTS FOR A SINGLE FOLDER	*
'This file will prompt a user for an ID and SERVER and create a hidden share.		*
'	It sets the default share permission of Everyone at Full control, and 			*
'  sets the user limit to the Maximum connections.							*
'
' 03/17/03 - change to add Archive folder setup 								*
'
'	It then sets the NTFS rights for a user with the same name as the folder		*
'  Assumes:															*
'	1) platform = WINDOWS 2000	(client and server)						*
'	2) SERVER >> User shares are created on the server under D:\DATA\USERS	*
'	3) Domain = DEN and User ID's match the share folder name				*
'	4) need to have file AdsSecurity.dll  in the same directory as this script.		*
'  Errors are reported in the LOGFILE variable (see below)						*
'									edited 03/17/03 by ncjws			*
'*************************************************************************************	''
Option Explicit
On Error Resume Next

Dim WshShell, WINDIR, fso
	set WshShell = WScript.CreateObject("WScript.Shell")
	WINDIR = WshShell.ExpandEnvironmentStrings("%WinDir%")
	set fso = CreateObject("Scripting.FileSystemObject")

	If not (fso.FileExists(WINDIR &"\system32\ADsSecurity.dll")) Then
		If fso.FileExists("ADsSecurity.dll") Then
			wscript.echo "File"& vbcrlf &"ADsSecurity.dll"& vbcrlf &"Doesn't exists so copying then registering it for this script to work"
			fso.CopyFile "ADsSecurity.dll", WINDIR &"\system32\", true
			WshShell.run "regsvr32 "& WINDIR &"\system32\ADsSecurity.dll", 1, true
		Else
			wscript.echo "File"& vbcrlf &"ADsSecurity.dll"& vbcrlf &"Does not Exist to copy"& vbcrlf &"Please copy this DLL to the same directory as this script"& vbcrlf &"before continuing."
			set WshShell = nothing
			set WINDIR = nothing
			set fso = nothing
			wscript.quit
		End If
	End If

Dim strComputer, objArgs, I

Set objArgs = WScript.Arguments
For I = 0 to objArgs.Count - 1
	If objArgs(I) = "/?" then
		Call ShowUsage()
	End If
Next

Dim WshNetwork
	set WshNetwork = WScript.CreateObject("WScript.Network")
Dim UserName1
	UserName1 = Wshnetwork.UserName

Dim THISBOX
	THISBOX = "DEN-"

'Log file for time and errors ''
DIM LOGFILE
	LOGFILE = "User-CreateNewHDLog.txt"
Dim f, FILE, EndMsg

Dim objWMIService, objNewShare, DIRECTORY, UserShare, UserFolder
Const FILE_SHARE = 0
Const NUMBERALLOWED = nothing

Dim UserArchiveFolder, XARCHIVE


'Begin Repeat Sequence
	Dim Response

Do Until Response = vbNo
		f = ""
Dim SINGLESHARE, USERNAME
	USERNAME = ""
Dim SHARENAME, SHAREDESC

If objArgs.Count = "2" then 
	USERNAME = WScript.Arguments.Unnamed.Item(1)
	strComputer = WScript.Arguments.Unnamed.Item(0)
	Dim QuickShare
	QuickShare = MsgBox("Ready to create share on "& vbcrlf & vbcrlf & strComputer & vbcrlf & vbcrlf &"     with user ID of "& vbcrlf & vbcrlf & USERNAME & vbcrlf & vbcrlf &"CONTINUE?" , vbYesNo, "Single User share?")
	If QuickShare = 7 then Call ExitScript()
End If

If objArgs.Count = "1" then 
	strComputer = WScript.Arguments.Unnamed.Item(0)
	SINGLESHARE = MsgBox("Do you wish to create a single User folder share, also creating a new folder for them?", vbYesNo, "Single User share?")
	If SINGLESHARE = 6 then  'YES
		USERNAME = InputBox("Enter the USER ID"& vbcrlf & vbcrlf &"(Ensure the ID has been created on then DEN domain first)", "user1")
		If USERNAME = "" then Call ExitScript()
	End If
	If SINGLESHARE = 7 then  'NO
		If strComputer = "" then Call ExitScript()
	End If
End IF

If objArgs.Count = "0" then
	USERNAME = InputBox("Enter the USER ID"& vbcrlf & vbcrlf &"(Ensure the ID has been created on the DEN domain first)", "Home Directory/Archive Folder Creation")
	If USERNAME = "" then 
		Call ExitScript()
	End If
	strComputer = InputBox("On which Server do you want to create the user's Home Directory?"& vbcrlf & vbcrlf & vbcrlf & vbcrlf &"ie: DEN-USER1", "ENTER SERVER NAME", THISBOX)
	If strComputer = "" then 
		Call ExitScript()
	End If
	If strComputer = "DEN-" then 
		Call ExitScript()
	End If
End If

	'Log file for time and errors ''
	EndMsg = ""	

	'Set objects to create the share ''
		Err.Clear
		Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
		If not Err.Number = 0 Then
			f = f & vbcrlf &  now &"  Unable to connect to server "& strComputer &" ...Error "& Err.Number &"  "& Err.Description
			f = f & vbcrlf
			wscript.echo " Server not found...check log file"
			Call EnterLogFile()
			Call ExitScript()
		End If
		Set objNewShare = objWMIService.Get("Win32_Share")
	
	'Setup User folder share names and properties for the share ''
	strComputer = UCASE(strComputer)
	SELECT CASE strComputer
			Case "DEN-USER1"
				set UserShare = fso.GetFolder("\\"& strComputer &"\D$\data1\Users1")
			DIRECTORY = UserShare		
			Case "DEN-USER2"
				set UserShare = fso.GetFolder("\\"& strComputer &"\D$\data2\Users2")
			DIRECTORY = UserShare		
			Case "DEN-USER3"
				set UserShare = fso.GetFolder("\\"& strComputer &"\E$\data3\Users3")
			DIRECTORY = UserShare		
			Case "DEN-USER4"
				set UserShare = fso.GetFolder("\\"& strComputer &"\E$\Data4\Users4")
			DIRECTORY = UserShare		
			Case "DEN-D"
				set UserShare = fso.GetFolder("\\"& strComputer &"\D$\data\Users")
			DIRECTORY = UserShare		
			Case "DEN-DEV"
				set UserShare = fso.GetFolder("\\"& strComputer &"\D$\data\Users")
			DIRECTORY = UserShare		
		End Select
		
	strComputer = UCASE(strComputer)
	SELECT CASE strComputer
		Case "DEN-USER1"
			set XArchive = fso.GetFolder("\\DEN-ARCHIVE\XARCHIVEe\A1")		
		Case "DEN-USER2"
			set XArchive = fso.GetFolder("\\DEN-ARCHIVE\XARCHIVEd\A2")		
		Case "DEN-USER3"
			set XArchive = fso.GetFolder("\\DEN-ARCHIVE\XARCHIVEe\A3")		
		Case "DEN-USER4"
			set XArchive = fso.GetFolder("\\DEN-ARCHIVE\XARCHIVEd\A4")		
		Case "DEN-D"
			set XArchive = fso.GetFolder("\\DEN-ARCHIVE\XARCHIVEe\D")		
		Case "DEN-DEV"
			set XArchive = fso.GetFolder("\\DEN-ARCHIVE\XARCHIVEd\DEV")		
		Case Else
			set XArchive = fso.GetFolder("\\DEN-ARCHIVE\XARCHIVE")		
	End Select
	

	'Create single share
	Call CREATESingleSHARE()

	'Append or create LOG FILE (appends to the beginning)
	Call EnterLogFile()

	'Repeat if necessary	
	Response = MsgBox (EndMsg & vbcrlf &" Check "& LOGFILE &" for details"& vbcrlf & vbcrlf &"Do you want to create another Home Directory/Archive Folder?", 36, "Complete...Create another share?")
	
	WshShell.LogEvent 0, Username1 &"... ran the User-CreateNewHD.vbe"& vbcrlf &" for user "& USERNAME & vbcrlf &" results were: "& EndMsg, strComputer
	USERNAME = ""
	strComputer = ""
Loop

'Close log file and exit ''
Call ExitScript()



'*************************************************************
'Sub routines to create Single share and append log file ''
'*************************************************************
SUB CREATESingleSHARE()
	On Error Resume Next
	'Check for user ID before creating shares or adding rights	
	Dim objUser, strUserName
	objUser = ""
	strUserName = USERNAME
	Err.Clear
	
	set objUser = GetObject("WinNT://DEN/" & strUserName)
	
	If not Err.Number = 0 Then
		f = f & vbcrlf &  now &" - "& USERNAME &" --- User ID not found in DEN domain...no share/archive folder created + NTFS not applied, error number: "& Err.Number &" "& Err.Description
			EndMsg = "User ID '"& USERNAME &"' not found in DEN domain"
		exit sub
	End If
	
	If not (fso.FolderExists("\\"& strComputer &"\D$\data\Users\"& USERNAME)) Then
		UserFolder = fso.CreateFolder("\\"& strComputer &"\D$\data\Users\"& USERNAME)
	End If
		set UserFolder = fso.GetFolder("\\"& strComputer &"\D$\data\Users\"& USERNAME)

	Dim errReturn
	SHARENAME = UserFolder.Name & Chr(36)
	SHAREDESC = UserFolder.Name &"'s Home directory"
	USERNAME = UserFolder.Name
	errReturn = objNewShare.Create ("D:\data\Users\"& UserFolder.Name, SHARENAME, FILE_SHARE, NUMBERALLOWED, SHAREDESC)
		If errReturn = 9 then 
			EndMsg = "Share FAILED for user '"& USERNAME &"'...Invalid share Name"
			f = f & vbcrlf &  now &" - "& UserFolder.Name &" --- share FAILED...Invalid share Name, error number: "& errReturn &" "& Err.Description
			Exit Sub
		End If
		If errReturn = 22 then 
			EndMsg = "Share FAILED for '"& USERNAME &"'...may already exist"
			f = f & vbcrlf &  now &" - share FAILED on '"& UserFolder.Name &"'...may already exist, error number: "& errReturn &" "& Err.Description 
			Exit Sub
		End If
		If errReturn = 24 then 
			EndMsg = "Share FAILED for '"& USERNAME &"'...check syntax"
			f = f & vbcrlf &  now &" - share FAILED on '"& UserFolder.Name &"'...check syntax, error number: "& errReturn &" "& Err.Description
			Exit Sub
		End If
		If errReturn = 0 then
			EndMsg = "Home Directory for '"& USERNAME &"' was successfully created on '"& strComputer &"'."
			f = f & vbcrlf &  now &" - Home Directory created successfully for '"& UserFolder.Name &"' - by "& Username1 &" on server "& strComputer
			Call ADDRIGHTS(DIRECTORY &"\"& UserFolder.Name, "DEN\"& USERNAME)

			'Create Archive Folder
			Call CREATEArchiveFolder()
		End If
END SUB

'Setup the archive folder for the user on DEN-ARCHIVE
SUB CREATEArchiveFolder()
	If not (fso.FolderExists(XArchive &"\"& USERNAME)) Then
		UserArchiveFolder = fso.CreateFolder(XArchive &"\"& USERNAME)
			If Err.Number = 0 then
				EndMsg = EndMsg & vbcrlf &"Archive Directory for '"& USERNAME &"' was successfully created."
				f = f & vbcrlf &  now &" - Archive Directory created successfully for '"& USERNAME &"' - by "& Username1
				Call ADDRIGHTS(UserArchiveFolder, "DEN\"& USERNAME)
				Call CreateArchiveSHARE() 
			Else
				EndMsg = EndMsg & vbcrlf &"ERROR creating Archive Directory for '"& USERNAME &" - Error = "& Err.number &"-"& Err.Description
				f = f & vbcrlf &  now &"ERROR creating Archive Directory for '"& USERNAME &" - Error = "& Err.number &"-"& Err.Description
			End If
	Else
		EndMsg = EndMsg & vbcrlf &"** Archive Directory for '"& USERNAME &"' already exists**."
		f = f & vbcrlf &  now &" - **Archive Directory for '"& USERNAME &"' already exists**."
	End If

END SUB

'*************************************************************
'Sub routines to create Single share and append log file ''
'*************************************************************
SUB CreateArchiveSHARE()
	On Error Resume Next
	Dim objWMIServiceX, objNewShareX
	
	Set objWMIServiceX = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\DEN-ARCHIVE\root\cimv2")
	If not Err.Number = 0 Then
		f = f & vbcrlf &  now &"  Unable to connect to server DEN-ARCHIVE ...Error "& Err.Number &"  "& Err.Description
		f = f & vbcrlf
		wscript.echo " Server not found...check log file"& VBCRLF &"LINE 234"
		Call EnterLogFile()
		Call ExitScript()
	End If
	Set objNewShareX = objWMIServiceX.Get("Win32_Share")

	'Check for user ID before creating shares or adding rights	
	Dim objUser, strUserName
	objUser = ""
	strUserName = USERNAME
	Err.Clear
	
	set objUser = GetObject("WinNT://DEN/" & strUserName)
	
	If not Err.Number = 0 Then
		f = f & vbcrlf &  now &" - "& USERNAME &" --- User ID not found in DEN domain...no share/archive folder created + NTFS not applied, error number: "& Err.Number &" "& Err.Description
			EndMsg = "User ID '"& USERNAME &"' not found in DEN domain"
		exit sub
	End If
	
	
	'If not (fso.FolderExists("\\DEN-ARCHIVE\E$\"& USERNAME)) Then
	'	UserFolder = fso.CreateFolder("\\DEN-ARCHIVE\E$\"& USERNAME)
	'End If
		'set UserFolder = fso.GetFolder("\\DEN-ARCHIVE\E$\"& USERNAME)
		set UserFolder = fso.GetFolder(UserArchiveFolder)
		
	Dim errReturn
	SHARENAME = UserFolder.Name & Chr(36)
	SHAREDESC = UserFolder.Name &"'s Archive directory"
	USERNAME = UserFolder.Name
	
	'errReturn = objNewShareX.Create ("E:\"& UserFolder.Name, SHARENAME, FILE_SHARE, NUMBERALLOWED, SHAREDESC)
	
	errReturn = objNewShareX.Create ("E:\" & UserFolder.ParentFolder.Name &"\"& UserFolder.Name, SHARENAME, FILE_SHARE, NUMBERALLOWED, SHAREDESC)
	
		If errReturn = 9 then 
			EndMsg = "Archive Share FAILED for user '"& USERNAME &"'...Invalid share Name"
			f = f & vbcrlf &  now &" - "& UserFolder.Name &" --- Archive share FAILED...Invalid share Name, error number: "& errReturn &" "& Err.Description
			Exit Sub
		End If
		If errReturn = 22 then 
			EndMsg = "Archive Share FAILED for '"& USERNAME &"'...may already exist"
			f = f & vbcrlf &  now &" - Archive share FAILED on '"& UserFolder.Name &"'...may already exist, error number: "& errReturn &" "& Err.Description 
			Exit Sub
		End If
		If errReturn = 24 then 
			EndMsg = "Archive Share FAILED for '"& USERNAME &"'...check syntax"
			f = f & vbcrlf &  now &" - Archive share FAILED on '"& UserFolder.Name &"'...check syntax, error number: "& errReturn &" "& Err.Description
			Exit Sub
		End If
		If errReturn = 0 then
			EndMsg = "Archive Directory for '"& USERNAME &"' was successfully created on 'DEN-ARCHIVE'."
			f = f & vbcrlf &  now &" - Archive Directory created successfully for '"& UserFolder.Name &"' - by "& Username1 &" on server DEN-ARCHIVE"
		End If
END SUB


'********************************************************************************************
'*** Access Control Entry Inheritance Flags
'*** Possible values for the IADsAccessControlEntry::AceFlags property.  
	const ADS_ACEFLAG_UNKNOWN					= &h1
'*** child objects will inherit ACE of current object
	const ADS_ACEFLAG_INHERIT_ACE				= &h2	
'*** prevents ACE inherited by the object from further propagation
	const ADS_ACEFLAG_NO_PROPAGATE_INHERIT_ACE	= &h4
'*** indicates ACE used only for inheritance (it does not affect permissions on object itself)
	const ADS_ACEFLAG_INHERIT_ONLY_ACE 			= &h8
'*** indicates that ACE was inherited	
	const ADS_ACEFLAG_INHERITED_ACE				= &h10
'*** indicates that inherit flags are valid (provides confirmation of valid settings)
	const ADS_ACEFLAG_VALID_INHERIT_FLAGS		= &h1f
'*** for auditing success in system audit ACE
	const ADS_ACEFLAG_SUCCESSFUL_ACCESS		= &h40
'*** for auditing failure in system audit ACE
	const ADS_ACEFLAG_FAILED_ACCESS				= &h80
'*** Access Control Entry Type Values
'*** Possible values for the IADsAccessContronEntry::AceType property.  
	const ADS_ACETYPE_ACCESS_ALLOWED			= 0
	const ADS_ACETYPE_ACCESS_DENIED			= &h1
	const ADS_ACETYPE_SYSTEM_AUDIT				= &h2
	const ADS_ACETYPE_ACCESS_ALLOWED_OBJECT    	= &h5
	const ADS_ACETYPE_ACCESS_DENIED_OBJECT    	= &h6
	const ADS_ACETYPE_SYSTEM_AUDIT_OBJECT     	= &h7
'********************************************************************************************

'SUB TO ADD NTFS RIGHTS TO EMPTY FOLDER
SUB ADDRIGHTS(FOLDER, USER)
'On Error Resume Next
Dim fldr, sf, f1
'* MAIN AREA ************************************
	Dim sec, sd, dacl, ace
	Set sec = CreateObject("ADsSecurity")
	Set sd = sec.GetSecurityDescriptor(CStr("FILE://"& FOLDER))
	Set dacl = sd.DiscretionaryAcl
	Set ace = CreateObject("AccessControlEntry")
	ace.AceType = "0"
	ace.AccessMask = "1245631"
	ace.trustee = USER
	ace.AceFlags = "3"
	ace.Flags = "0"
'*** Apply the rights now  ''
	dacl.AddACE ace    
	Call ReorderDACL(dacl)
	sd.DiscretionaryAcl = dacl
	sec.SetSecurityDescriptor sd
'*Clean ups the subs ''
Call CleanupSubFldrs(FOLDER, USER)
Call CleanupSubFiles(FOLDER, USER)
f = f & vbcrlf &  " -- NTFS rights applied to Folder "& FOLDER
END SUB


'*** Subroutine reordering the ACLs (per Q279682)
'*** ACEs need to be properly ordered, since AddAce method does not perform ordering. 
'*** If an access-allowed ACE appears before access-denied, a trustee will be granted access.
'*** The preferred order of ACEs in a DACL is described in MSDN Library (at msdn.microsoft.com). 
'*** For Windows 2000, ACEs should be arranged into two main groups - non-inherited and inherited.
'*** Non-inherited ACEs should be listed first, followed by the inherited ones. Within each group
'*** (non-inherited and inherited), ACEs are arranged in the following fashion:
'*** - access-denied ACEs that apply to the object itself
'*** - access-denied ACEs that apply to subobjects of the object (including its properties)
'*** - access-allowed ACEs that apply to the object itself
'*** - access-allowed ACEs that apply to subobjects of the object (including its properties) 
'*** Since the script does not affect inherited ACEs (it sets permission directly on target object)
'*** they do not have to be rearranged. We only need to rearrange non-inherited ACEs
Sub ReorderDACL(odacl)
	Dim oNewDACL		'object used to temporarily store DACL (during ordering)
	Dim oInheritedDACL	'object representing list of all Inherited ACEs
	Dim oDenyDACL		'object representing list of non-Inherited Deny ACEs
	Dim oAllowDACL		'object representing list of non-Inherited Allow ACEs
	Dim oACE			'object representing ACE (used for enumeration)
'*** Create Access Control List objects 
	Set oNewDACL = CreateObject("AccessControlList")
	Set oInheritedDACL = CreateObject("AccessControlList")
	Set oAllowDACL = CreateObject("AccessControlList")
	Set oDenyDACL = CreateObject("AccessControlList")
'*** Add individual ACEs into each of the lists
'*** based on the ACE Flags and ACE Type values
	For Each oACE In oDACL 
	If ((oACE.AceFlags AND ADS_ACEFLAG_INHERITED_ACE) = ADS_ACEFLAG_INHERITED_ACE) Then	 
'*** as explained, no sorting is needed for Inherited ACEs, they are simply
'*** added to the list and retrieved at the end of the sub in the same order
		oInheritedDACL.AddAce oACE 
	Else
'*** non-Inherited ACEs need to be placed in their respective list to be re-ordered
		Select Case oACE.AceType	 
			Case ADS_ACETYPE_ACCESS_ALLOWED		 
				oAllowDACL.AddAce oACE	 
			Case ADS_ACETYPE_ACCESS_DENIED		 
				oDenyDACL.AddAce oACE    
		End Select
	End If
	Next
'**************************************************
'*** Recreate the Access Control List following the appropriate order
'*** - non-Inherited Deny ACEs
'*** - non-Inherited Allow ACEs
'*** - Inherited ACEs
	For Each oACE In oDenyDACL 
		 oNewDACL.AddAce oACE 
	Next
	For Each oACE In oAllowDACL
		  oNewDACL.AddAce oACE
	Next 
	For Each oACE In oInheritedDACL
		oNewDACL.AddAce oACE 
	Next
	Set oInheritedDACL = Nothing
	Set oDenyDACL = Nothing
	Set oAllowDACL = Nothing
'**************************************************
'*** Set appropriate DACL revision level
	oNewDACL.AclRevision = oDACL.AclRevision
'**************************************************
'*** Reset the original DACL
	Set oDACL = Nothing
	Set oDACL = oNewDACL
End Sub

'Cleanup Inheritance of Subs
SUB CleanupSubFldrs(SubFolder, USER)
Dim sd, dacl, ace, sec, SSubFolder
	Set sec = CreateObject("ADsSecurity")
	For each SSubFolder in (FSO.GETFOLDER(SubFolder)).SubFolders
		Set sd = sec.GetSecurityDescriptor(CStr("FILE://"& SSubFolder))
		Set dacl = sd.DiscretionaryAcl
		Set ace = CreateObject("AccessControlEntry")
		ace.trustee = USER
		ace.AccessMask = "1245631"
		ace.AceType = "0"
		ace.AceFlags = "19"
		ace.Flags = "0"
		dacl.AddACE ace 
		sd.DiscretionaryAcl = dacl
		sec.SetSecurityDescriptor sd
		Call CleanupSubFiles(SSubFolder, USER)
		call CleanupSubFldrs(SSubFolder, USER)
	Next
End Sub

SUB CleanupSubFiles(SubFolder, USER)
Dim sd, dacl, ace, sec, SubFile
	Set sec = CreateObject("ADsSecurity")
	For each SubFile in (FSO.GETFOLDER(SubFolder)).Files
		Set sd = sec.GetSecurityDescriptor(CStr("FILE://"& SubFile))
		Set dacl = sd.DiscretionaryAcl
		Set ace = CreateObject("AccessControlEntry")
		ace.trustee = USER
		ace.AccessMask = "1245631"
		ace.AceType = "0"
		ace.AceFlags = "16"
		ace.Flags = "0"
		dacl.AddACE ace 
		sd.DiscretionaryAcl = dacl
		sec.SetSecurityDescriptor sd
	Next
End Sub

SUB ShowUsage
	Wscript.echo "'************* CREATE SHARE AND APPLY NTFS RIGHTS FOR A SINGLE FOLDER	"& vbcrlf &_
		"This file will prompt a user for an ID and SERVER and create a hidden share.		"& vbcrlf &_
		"	It sets the default share permission of Everyone at Full control, and 			"& vbcrlf &_
		"  sets the user limit to the Maximum connections.							"& vbcrlf &_
		"	It then sets the NTFS rights for a user with the same name as the folder		"& vbcrlf &_
		"  Assumes:															"& vbcrlf &_
		"	1) platform = WINDOWS 2000	(client and server)						"& vbcrlf &_
		"	2) SERVER >> User shares are created on the server under D:\DATA\USERS	"& vbcrlf &_
		"	3) Domain = DEN and User ID's match the share folder name				"& vbcrlf &_
		"	4) need to have file AdsSecurity.dll  in the same directory as this script.		"& vbcrlf & vbcrlf &_
		"  Errors are reported in the LOGFILE variable (see below)						"& vbcrlf &_
		"									created 10/02/02 by ncjws			"& vbcrlf &_
		"*************************************************************************************	"& vbcrlf & vbcrlf &_
		"    COMMAND LINE SYNTAX   =     NewUserShare+Rights.vbs SERVER USERID        ...(1 space between each)  "
	set WshShell = nothing
	set WINDIR = nothing
	set fso = nothing
	set strComputer = nothing 
	set objArgs = nothing 
	set I = nothing
	wscript.quit
End Sub

SUB EnterLogFile()
	Dim LogFileText, ReadOldLogFile
	ReadOldLogFile = ""
	If fso.FileExists(LOGFILE) then
		Set LogFileText= fso.OpenTextFile(LOGFILE, 1, True)
		ReadOldLogFile = LogFileText.ReadAll
		LogFileText.Close
	End If
	Set LogFileText= fso.CreateTextFile(LOGFILE, True)
	LogFileText.WriteLine(f & vbcrlf & ReadOldLogFile)
	LogFileText.Close
	Set LogFileText = nothing
	Set ReadOldLogFile = nothing
End Sub

SUB ExitScript()
	On Error Resume Next
	Set strComputer = nothing
	Set I = nothing
	Set WshShell = nothing
	Set WINDIR = nothing
	Set WshNetwork = nothing
	Set UserName1 = nothing
	Set THISBOX = nothing
	Set EndMsg = nothing
	Set objWMIService = nothing
	Set objNewShare = nothing
	Set DIRECTORY = nothing
	Set UserShare = nothing
	Set UserArchive = nothing
	Set UserFolder = nothing
	Set Response = nothing
	Set SINGLESHARE = nothing
	Set USERNAME = nothing
	Set SHARENAME = nothing
	Set SHAREDESC = nothing
	Set Quickshare = nothing

	Set LOGFILE = nothing
	Set FILE = nothing 
	Set f = nothing
	Set fso = nothing
	wscript.quit
End SUB