'-------------------------------------------------------
'
'  Sample script that creates a network drive mapping
'  to the next available drive letter on the host
'
'  Shawn Stugart
'=======================================================

set oFSO = CreateObject("Scripting.FileSystemObject")
drive = "C"

Do While oFSO.DriveExists(drive & ":")
    iChar = Asc(drive)
    drive = Chr(ichar + 1)
Loop 

set oNetwork = CreateObject("Wscript.Network")
oNetwork.MapNetworkDrive drive &":", "\\london\labshare"


