Randomize
iNumber = Round(Rnd * 100)
iTries = 0

Do
	iTries = iTries + 1
	iGuess = InputBox("What number am I thinking of?")
	If iGuess = "" Then WScript.Quit
	iGuess = CInt(iGuess)
Loop Until Guesser(iGuess) = True

Function Guesser(iGuess)
	If iGuess > iNumber Then
		WScript.Echo iGuess & " is too high."
		Guesser = False
	Elseif iGuess < iNumber Then
		WScript.Echo iGuess & " is too low."
		Guesser = False
	Else
		WScript.Echo "You got it! The number was " & iNumber & ".  It took you " & iTries & " tries."
		Guesser = True
	End If
End Function