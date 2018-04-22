'*****************************************************************************
'										
'  Script to search a text file ("c:\clients.txt" by default) for computer   
'  names and look for games running on those computers.  If games are found  
'  running, you are prompted whether or not to terminate games on individual 
'  machines or to terminate ALL games found on all computers.			
'  										
'  Last Modified 4-13-2004
'  Shawn Stugart							
'*****************************************************************************






Option Explicit
On Error Resume Next
Dim hostname, sFileName, numGames, oFSO, oShell, oInputFile
Dim strComputer, objScriptExec, strPingStdOut, objWMIService
Dim colProcesses, proc, strResult, strFormatting, strFinal, userFile
Dim colSolProcesses, colFCProcesses, colWMProcesses, colPBProcesses
Dim strKillResponse, strWhich2Kill, strTryAgain, oFSO2, strGamers
Dim sComp, strKillAllOnBox, strTryAgain2, objWMI2, colProcessList
Dim objProcess, totalGames, colIEProcesses
const ForReading = 1
Const ForWriting = 2

Call CheckCScript

sFileName = "Script\clients.txt"
numGames = 0
set oFSO = CreateObject("Scripting.FileSystemObject")
set oShell = CreateObject("WScript.Shell")
set oInputFile = oFSO.OpenTextFile("c:\" & sFileName, ForReading)
If Err.Number <> 0 Then
    WScript.Echo "Couldn't Find C:\" & sFileName
    WScript.Quit
End If
Set userFile = oFSO.CreateTextFile("c:\Gamers.txt", True)

Do Until oInputFile.AtEndOfStream
    strComputer = oInputFile.Readline
    Set objScriptExec = oShell.Exec("%comspec% /c ping -n 1 -w 500 " & strComputer)
    strPingStdOut = Lcase(objScriptExec.StdOut.ReadAll)
    If InStr(strPingStdOut, "reply from ") Then
	    Set objWMIService = GetObject("winmgmts:" _
	        & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	    Select Case Err.Number
	        Case 0
			    Set colProcesses = objWMIService.ExecQuery("SELECT Name FROM Win32_Process")
			    For Each proc In colProcesses
			        If proc.Name = "sol.exe" Or proc.Name = "freecell.exe" Or _
			                proc.Name = "winmine.exe" Or proc.Name = "PINBALL.EXE" Or _
                            UCase(proc.Name) = "IEXPLORE.EXE" Then
			            strResult = GetGames()
			            numGames = numGames + 1
			        End If
			    Next
            Case 70
	            strFormatting = "--------------------------------------------" & vbCrLf _
		                    & strComputer & vbCrLf _
		                  & "............" & vbCrLf
			    strResult = "Permission to query this machine has been denied." & vbCrLf
			    strFinal = strFinal & strFormatting & strResult
	        Case Else
	            strFormatting = "--------------------------------------------" & vbCrLf _
		                    & strComputer & vbCrLf _
		                  & "............" & vbCrLf
			    strResult = "An unspecified error has occurred." & vbCrLf
			    strFinal = strFinal & strFormatting & strResult
	    End Select
	    strFormatting = "--------------------------------------------" & vbCrLf _
	                    & strComputer & vbCrLf _
	                  & "............" & vbCrLf
		If numGames <> 0 Then
		    userFile.WriteLine(strComputer)
		    strFinal = strFinal & strFormatting & strResult
		ElseIf Err.Number <> 0 Then
		    Err.Clear
		Else
		    strFinal = strFinal & strFormatting & "There do not appear to be any games running." & vbCrLf
		End If
                totalGames = totalGames + numGames
		numGames = 0
		strResult = Null
	Else
	    strFormatting = "--------------------------------------------" & vbCrLf _
	                    & strComputer & vbCrLf _
	                  & "............" & vbCrLf
	    strResult = strComputer & " Could Not Be Contacted" & vbCrLf
	    strFinal = strFinal & strFormatting & strResult
	End If
Loop
userFile.Close
WScript.Echo strFinal
WScript.StdOut.WriteLine("---------------------------------------------")
WScript.StdOut.WriteLine
WScript.StdOut.WriteLine "Total Games Running = " & totalGames
If totalGames = 0 Then WScript.Quit
WScript.StdOut.Write("Do you want to kill any of these games?[y/n]:")
strKillResponse = WScript.StdIn.ReadLine
WScript.StdOut.WriteLine
If strKillResponse = LCase("y") Then
    Do
	    WScript.StdOut.Write("Kill games on Individual machines or All?[i/a]:")
	    strWhich2Kill = WScript.StdIn.ReadLine
	    WScript.StdOut.WriteLine
	    Select Case LCase(strWhich2Kill)
	        Case "i"
	            Call OneByOne()
	            strTryAgain = "n"
	        Case "a"
	            Call KillAll()
	            strTryAgain = "n"
	        Case Else
	            WScript.StdOut.Write("You specified an invalid option.  Try again?[y/n]:")
	            strTryAgain = WScript.StdIn.ReadLine
	            WScript.StdOut.WriteLine
	        End Select
     Loop Until strTryAgain <> LCase("y")      
Else
    WScript.StdOut.WriteLine
    WScript.Echo "Thanks for playing!"
    WScript.Quit
End If

'******************************************************************************************************
Sub CheckCScript
	hostname = lcase(right(WSCript.Fullname, 11))
	If hostname = "wscript.exe" Then
	    WScript.Echo "This script requires cscript.exe"
	    WScript.Quit
	End If
End Sub
'******************************************************************************************************
Function GetGames()
    Set colSolProcesses = objWMIService.ExecQuery("SELECT Name FROM Win32_Process WHERE Name = 'sol.exe'")
    If colSolProcesses.Count <> 0 Then
        GetGames = GetGames & space(12) & "Solitaire IS running" & vbCrLf
    End If
    Set colFCProcesses = objWMIService.ExecQuery("SELECT Name FROM Win32_Process WHERE Name = 'freecell.exe'")
    If colFCProcesses.Count <> 0 Then
        GetGames = GetGames & space(12) & "Freecell IS running" & vbCrLf
    End If
    Set colWMProcesses = objWMIService.ExecQuery("SELECT Name FROM Win32_Process WHERE Name = 'winmine.exe'")
    If colWMProcesses.Count <> 0 Then
        GetGames = GetGames & space(12) & "WinMine IS running" & vbCrLf
    End If
    Set colPBProcesses = objWMIService.ExecQuery("SELECT Name FROM Win32_Process WHERE Name = 'PINBALL.EXE'")
    If colPBProcesses.Count <> 0 Then
        GetGames = GetGames & space(12) & "Pinball IS running" & vbCrLf
    End If
    Set colIEProcesses = objWMIService.ExecQuery("SELECT Name FROM Win32_Process WHERE Name = 'IEXPLORE.EXE'")
    If colIEProcesses.Count <> 0 Then
        GetGames = GetGames & space(12) & "IE IS running" & vbCrLf
    End If
End Function
'******************************************************************************************************
Sub OneByOne()
    Const ForReading = 1
    Set oFSO2 = CreateObject("Scripting.FileSystemObject")
    set strGamers = oFSO2.OpenTextFile("c:\Gamers.txt",ForReading)
    Do Until strGamers.AtEndOfStream
        sComp = strGamers.ReadLine
        Do
	        strTryAgain2 = "n"
	        WScript.StdOut.Write("Kill all games on " & sComp & "?[y/n]:")
	        strKillAllOnBox = WScript.StdIn.ReadLine
	        WScript.StdOut.WriteLine
	        Select Case LCase(strKillAllOnBox)
		        Case "y"
		            strTryAgain2 = "n"
			        Set objWMI2 = GetObject("winmgmts:" _
					    & "{impersonationLevel=impersonate}!\\" & sComp & "\root\cimv2")
					    If Err.Number = 0 Then
							Set colProcessList = objWMI2.ExecQuery _
							    ("SELECT Name FROM Win32_Process WHERE Name = 'sol.exe'")
							If Err.Number = 0 Then
								For Each objProcess in colProcessList
								    objProcess.Terminate()
								Next
							Else
							    Err.Clear
							End If
							Set colProcessList = objWMI2.ExecQuery _
							    ("SELECT Name FROM Win32_Process WHERE Name = 'freecell.exe'")
							If Err.Number = 0 Then
								For Each objProcess in colProcessList
								    objProcess.Terminate()
								Next
							Else
							    Err.Clear
							End If
							Set colProcessList = objWMI2.ExecQuery _
							    ("SELECT Name FROM Win32_Process WHERE Name = 'winmine.exe'")
							If Err.Number = 0 Then
								For Each objProcess in colProcessList
								    objProcess.Terminate()
								Next
						    Else
						        Err.Clear
						    End If
							Set colProcessList = objWMI2.ExecQuery _
							    ("SELECT Name FROM Win32_Process WHERE Name = 'PINBALL.EXE'")
							If Err.Number = 0 Then
								For Each objProcess in colProcessList
								    objProcess.Terminate()
								Next
							Else
							    Err.Clear
							End If
							Set colProcessList = objWMI2.ExecQuery _
							    ("SELECT Name FROM Win32_Process WHERE Name = 'IEXPLORE.EXE'")
							If Err.Number = 0 Then
								For Each objProcess in colProcessList
								    objProcess.Terminate()
								Next
							Else
							    Err.Clear
							End If
						Else
						    WScript.StdOut.WriteLine("Error killing games on " & sComp)
						End If
			    Case "n"
			        WScript.StdOut.WriteLine
			        strTryAgain2 = "n"
			    Case Else
			        WScript.StdOut.WriteLine
			        WScript.StdOut.Write("You specified an invalid option.  Try again?[y/n]:")
			        strTryAgain2 = WScript.StdIn.ReadLine
			        WScript.StdOut.WriteLine
			End Select
	    Loop Until strTryAgain2 <> LCase("y") 
	Loop
	WScript.Echo vbCrLF & vbCrLf & "Complete."
End Sub
'******************************************************************************************************
Sub KillAll()
    Const ForReading = 1
    Set oFSO2 = CreateObject("Scripting.FileSystemObject")
    set strGamers = oFSO2.OpenTextFile("c:\Gamers.txt",ForReading)
    Do Until strGamers.AtEndOfStream
        sComp = strGamers.ReadLine
        Set objWMI2 = GetObject("winmgmts:" _
			& "{impersonationLevel=impersonate}!\\" & sComp & "\root\cimv2")
				If Err.Number = 0 Then
					Set colProcessList = objWMI2.ExecQuery _
					    ("SELECT Name FROM Win32_Process WHERE Name = 'sol.exe'")
			    	For Each objProcess in colProcessList
					    objProcess.Terminate()
					Next
					Set colProcessList = objWMI2.ExecQuery _
					    ("SELECT Name FROM Win32_Process WHERE Name = 'freecell.exe'")
					For Each objProcess in colProcessList
					    objProcess.Terminate()
					Next
					Set colProcessList = objWMI2.ExecQuery _
					    ("SELECT Name FROM Win32_Process WHERE Name = 'winmine.exe'")
					For Each objProcess in colProcessList
						objProcess.Terminate()
					Next
					Set colProcessList = objWMI2.ExecQuery _
					    ("SELECT Name FROM Win32_Process WHERE Name = 'PINBALL.EXE'")
					For Each objProcess in colProcessList
					    objProcess.Terminate()
					Next
					Set colProcessList = objWMI2.ExecQuery _
					    ("SELECT Name FROM Win32_Process WHERE Name = 'IEXPLORE.EXE'")
					For Each objProcess in colProcessList
					    objProcess.Terminate()
					Next
					WScript.StdOut.WriteLine(sComp & "--Games Killed!")
				Else
					WScript.StdOut.WriteLine("Error killing games on " & sComp)
				End If
	Loop
	WScript.Echo vbCrLF & vbCrLf & "Complete."
End Sub

'Sub Syntax()
