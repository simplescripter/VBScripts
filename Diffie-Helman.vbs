Option Explicit

Dim iG, iP, iXa, iXb
Dim iYa, iYb

iG = InputBox("Enter a prime number:")
iP = InputBox("Enter another prime number GREATER than " & iG & ":")

iXa = InputBox("Enter a random number. Keep this secret, Student1")
iXb = InputBox("Enter a random number. Keep this secret, Student2")

iYa = (iG^iXa) Mod iP
iYb = (iG^iXb) Mod iP

WScript.echo "Student1 Ya = "  &  iYa & VbCrLf & VbCrLf & "Openly exchange this number " _
	& "with Student2."
WScript.echo "Student2 Yb = "  &  iYb & VbCrLf & VbCrLf & "Openly exchange this number " _
	& "with Student1."

wscript.echo "Now you can compute your secret key."
wscript.echo "Student1 key formula is (Yb^Xa) Mod P" & VbCrLf _
	& vbTab & vbTab & "= " & iYb & "^" & iXa & " Mod " & iP & VbCrLf _ 
	& vbTab & vbTab & "= " & (iYb^iXa) Mod iP & VbCrLf & VbCrLf _
	& "Student2 key formula is (Ya^Xb) Mod P" & VbCrLf _
	& vbTab & vbTab & "= " & iYa & "^" & iXb & " Mod " & iP _
	& VbCrLf & vbTab & vbTab & "= " & (iYa^iXb) Mod iP 
