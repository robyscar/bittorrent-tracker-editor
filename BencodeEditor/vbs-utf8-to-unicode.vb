' http://demon.tw/programming/vbs-utf8-to-unicode.html
Function Utf8ToUnicode(str)
	 Dim i, c, c2, c3, out, length
	out = ""
	i = 1
	length = LenB(str)
	Do  While i <= length
		c = AscB(MidB(str,i,1))
		i = i + 1
		Select  Case (c \ 2 ^ 4)
		     Case 0,1,2,3,4,5,6,7
		    	out = out & ChrW(c)
		    Case 12,13
		    	c2 = AscB(MidB(str,i,1))
		    	i = i + 1
		    	out = out & ChrW(((c And &H1F) * 2 ^ 6) Or (c2 And &H3F))
		     Case 14
		    	c2 = AscB(MidB(str,i,1))
		    	i = i + 1
		    	c3 = AscB(MidB(str,i,1))
		    	i = i + 1
		    	out = out & ChrW(((c And &H0F) * 2 ^ 12) Or _
		    			((c2 And &H3F) * 2 ^ 6) Or _
		    			((c3 And &H3F) * 2 ^ 0))
		 End  Select 
	Loop
	Utf8ToUnicode = out
End  Function