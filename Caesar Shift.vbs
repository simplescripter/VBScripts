Option Explicit

Dim sCipherText, i, j, sUnshiftedChar
Dim sResult

sCipherText = "DA VJG RTKEMKPI QH OA VJWODU UQOGVJKPI YKEMGF VJKU YCA EQOGU"
sCipherText = Normalize(sCipherText)

For j = 1 To 25
	For i = 1 To Len(sCipherText)
		sResult = sResult & ShiftIt(Mid(sCipherText, i, 1), j)
	Next
	WScript.Echo "Shift " & j & "  " & sResult
	sResult = ""
Next

Function Normalize(sToNormalize)
	Normalize = Replace(sToNormalize," ","")
End Function

Function ShiftIt(sCipherChar, iShift)
		sUnshiftedChar = Asc(sCipherChar)
		If sUnshiftedChar + iShift > 90 Then
			ShiftIt = Chr(((sUnshiftedChar + iShift) - 90) + 64)
		Else
			ShiftIt = Chr(sUnshiftedChar + iShift)
		End If
End Function