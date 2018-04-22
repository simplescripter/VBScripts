option explicit

dim i, j, maxlen, pwdval, dbg

randomize
maxlen = 8

'wscript.Echo asc("0") & asc("A") & asc("a")

for i = 1 to maxlen
   do
       j = round(rnd * 128)
   loop until (j >= 48 and j <= 57) or (j >= 65 and j <= 90) _
    or (j >= 97 and j <= 122)
   dbg = dbg & j & vbCrLf
   pwdval = pwdval & chr(j)
next

wscript.echo dbg & vbCrLf & pwdval
   
	