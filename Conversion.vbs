strInput = InputBox("Do you want to convert a temperature from: " _
    & vbCrLf & "1" & vbTab & "Farenheit to Celsius?" _
    & vbCrLf & "2" & vbTab & "Celsius to Farenheit?", "Enter a number.")

Select Case strInput
    Case "1"
        far = InputBox("Enter degrees Farenheit to convert:")
        WScript.Echo far & " degrees Farenheit is " _
        & Conv2Far(far) & " degrees Celsius."
    Case "2"
        cel = InputBox("Enter degrees Celsius to convert:")
			WScript.Echo cel & " degrees Celsius is " _
			& Conv2Cel(cel) & " degrees Farenheit."
    Case Else
	WScript.Echo strInput & " is an invalid value."
	WScript.Quit
End Select

Function Conv2Far(cel)
    cel = (far - 32) * .6
    Conv2Far = cel
End Function

Function Conv2Cel(far)
    far = (cel * 1.8) + 32
    Conv2Cel = far
End Function