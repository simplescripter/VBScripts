dim oNet, oNetDrives, i

set oNet = CreateObject("WScript.Network")
set oNetDrives = oNet.EnumNetworkDrives
For i = 0 to oNetDrives.Count - 1 Step 2
   oNet.RemoveNetworkDrive oNetDrives.Item(i)
Next

WScript.Echo "Finished Unmapping Network Drives."
