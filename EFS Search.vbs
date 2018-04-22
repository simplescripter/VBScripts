'---------------------------------------------------------------------------
'
'   This script loads a text file called c:\clients.txt and scans the
'   computers listed to find out if the user currently logged on to each
'   computer has enabled EFS.
'
'   Shawn Stugart
'===========================================================================


Option Explicit

dim strKeyPath, sFileName, oFSO, oShell, oInputFile
dim strComputer, objReg, arrSubKeys, subkey, sResult
dim key, j, objWMIService, colItems, objItem, strUserName
dim strEFSUsers, strNeverUsed

Const HKEY_CURRENT_USER = &H80000001
const ForReading = 1

strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
sFileName = "clients.txt"

set oFSO = CreateObject("Scripting.FileSystemObject")
set oShell = CreateObject("WScript.Shell")
set oInputFile = oFSO.OpenTextFile("c:\script\" & sFileName, ForReading)

On Error Resume Next

Do Until oInputFile.AtEndOfStream
    strComputer = oInputFile.Readline
    Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
        strComputer & "\root\default:StdRegProv")
    If Err.Number <> 0 Then
        sResult = sResult & "Couldn't read " & strComputer & vbCrLf
        Err.Clear
    Else
	    objReg.EnumKey HKEY_CURRENT_USER, strKeyPath, arrSubKeys
	    j = 0
	    For Each key in arrSubkeys
	        If key = "EFS" Then
	            j = 1
	            Exit For
	        End If    
	    Next
	    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
		Set colItems = objWMIService.ExecQuery("Select * from Win32_Environment")
		For Each objItem in colItems
		    strUserName =  objItem.UserName
		Next
		If j = 1 Then
		    strEFSUsers = strEFSUsers & "User " & strUserName & " HAS used EFS on "& UCase(strComputer) & vbCrLf
		Else   
		    strNeverUsed = strNeverUsed & "User " & strUserName & " has NEVER used EFS on " & UCase(strComputer) & vbCrLf
	    End If
	End If
Loop
WScript.Echo vbCrLf
WScript.Echo "Users who have enabled EFS on their computers:"
WScript.Echo "----------------------------------------------"
WScript.Echo strEFSUsers
WScript.Echo vbCrLf
WScript.Echo "Users who have never enabled EFS on their computers:"
WScript.Echo "----------------------------------------------------"
WScript.Echo strNeverUsed
WScript.Echo vbCrLf
WScript.Echo "----------------------------------------------"
WScript.Echo sResult