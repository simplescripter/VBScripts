'ResetPW.vbs 
'Copyright 2000, by Chris Brooke

Option Explicit 
Dim objContainer, colUsers, lFlag 
Set objContainer=GetObject("WinNT://domain")

ObjContainer.Filter=Array("User") 

For Each colUsers in objContainer 
   IFlag=colUsers.Get("UserFlags") 
   If (lFlag AND &H10000) <> 0 Then 
     ColUsers.Put "UserFlags", lFlag XOR &H10000 
   End If 
   colUsers.Put "PasswordExpired", 1 
   colUsers.SetInfo 
Next


