Option Explicit

Dim i, j, maxlen, pwdval, rndPassword, iValidation

Randomize
maxlen = 4

For i = 1 to maxlen
	Do
		j = round(rnd * 122)
	loop until (j >= 48 And j <= 122)
pwdval = pwdval & chr(j)
Next

rndPassword = InputBox("Enter the following code EXACTLY if you want to proceed:" & VbCrLf & VbCrLf & pwdval)
If rndPassword = pwdval Then
	iValidation = 1
Else
	iValidation = 0
End If

WScript.Echo iValidation

   
	