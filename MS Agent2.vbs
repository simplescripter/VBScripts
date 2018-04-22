' A script to demonstrate manipulating MS Agent characters.  This script
' enumerates the actions a paricular character can perform then terminates.
' For this script to work properly, you'll need to make sure you've got
' a text-to-speech engine installed, along with the Microsoft SAPI 4.0
' runtime support.  These requirements can be downloaded at 
' www.microsoft.com/msagent/downloads/user.asp
'
' Shawn Stugart
'
' 8/29/2005

strAgentName = InputBox("Enter a character name." & vbCrLf & "For example, " _
    & vbCrLf & vbCrLf & "Merlin" & vbCrLf & "Genie" & vbCrLf & "Robby" & vbCrLf & "Peedy")
strSpeech = InputBox("What would you like " & strAgentName & " to say?")
strAgentPath = "c:\windows\msagent\chars\" & strAgentName & ".acs"
Set objAgent = CreateObject("Agent.Control.2")

objAgent.Connected = TRUE
objAgent.Characters.Load strAgentName, strAgentPath
Set objCharacter = objAgent.Characters.Character(strAgentName)

objCharacter.Show
objCharacter.MoveTo 500,400

objCharacter.Speak strSpeech
WScript.Sleep 20000

'For Each strName in objCharacter.AnimationNames
'    If strName <> "Show" Then
'        If strName <> "Hide" Then
'	    objCharacter.Speak strName
'	    Set objRequest = objCharacter.Play(strName)
'	    i = 1
'           Do While objRequest.Status > 0
'		i = i + 1
'		WScript.Sleep 100
'		If i > 50 Then Exit Do
'	    Loop
'	    objCharacter.Stop
'	End If
'    End If
'Next


