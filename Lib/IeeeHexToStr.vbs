Function IeeeHexToStr(strHex)
	Dim i, dblFraction, lngHIFractBits, strHIFractBits, lngLOFractBits, strLOFractBits, strSign, intSign, strExp, lngExp, strMask, lngMask, strReturn
	Dim sngFraction, blnDouble, lngExpMask, intExpShift, intExpBias
	
	dblFraction = CDbl(0)
	lngHIFractBits = CLng(0)
	lngLOFractBits = CLng(0)
	strReturn = CStr("")
	
	'If a string of 16 characters is passed, then the function assumes double precision, if the string of 8 characters is passed then it assumes single precision
	If (Len(strHex) = 16) Then
		blnDouble = True
	Else
		blnDouble = False
	End If
	
	If blnDouble Then
		strHIFractBits = "&H" & Mid(strHex, 4, 5)
		lngHIFractBits = CLng(strHIFractBits) And &HFFFFF
	End If
	strLOFractBits = "&H" & Right(strHex, 8)
	lngLOFractBits = CLng(strLOFractBits)
	'If the fraction bits = 0 then assign the Return value = "0.0" and Exit
	If (blnDouble And (lngHIFractBits = 0) And (lngLOFractBits = 0)) Or ((Not blnDouble) And (lngLOFractBits = 0)) Then
		IeeeHexToStr = "0.0"
		Exit Function
	End If
	
	'Generate the Fractional value
	If blnDouble Then
		For i = 19 To 0 Step -1
			If BitStat(lngHIFractBits, i) Then
				dblFraction = dblFraction + (2 ^ -(20 - i))
			End If
		Next
		For i = 31 To 0 Step -1
			If BitStat(lngLOFractBits, i) Then
				dblFraction = dblFraction + (2 ^ -(52 - i))
			End If
		Next
		dblFraction = 1 + dblFraction
	Else
		For i = 22 To 0 Step -1
			If BitStat(lngLOFractBits, i) Then
				sngFraction = sngFraction + (2 ^ -(23 - i))
			End If
		Next
		sngFraction = 1 + sngFraction
	End If
	
	'Generate the Sign value
	strSign = "&H" & Left(strHex,1)
	intSign = CInt(strSign)
	If intSign > 7 Then
		intSign = -1
	Else
		intSign = 1
	End If

	'Generate the Exponent value
	strExp = "&h" & Left(strHex,3) & "00000"
	If blnDouble Then
		lngExpMask = &h7FF00000
		intExpShift = &h14
		intExpBias = &h3FF
	Else
		lngExpMask = &h7F800000
		intExpShift = &h17
		intExpBias = &h7F
	End If
	lngExp = CLng(strExp) And lngExpMask
	strMask = "&hFFFFFFFF"
	lngMask = Clng(strMask)
	For i = 1 to intExpShift
		lngMask = lngMask And Not (2^(intExpShift - 1))
	Next
	lngExp = (lngExp And lngMask) / (2 ^ intExpShift)
	lngExp = lngExp - intExpBias
	
	'Generate the Return string value
	If blnDouble Then
		dblFraction = (intSign) * (2 ^ lngExp) * (dblFraction)
		intValue = Fix(dblFraction)
		strReturn = CStr(dblFraction)
		If (dblFraction - intValue = 0.0 ) Then
			strReturn = strReturn & ".0"
		End If
	Else
		sngFraction = (intSign) * (2 ^ lngExp) * (sngFraction)
		intValue = Fix(sngFraction)
		strReturn = CStr(sngFraction)
		If (sngFraction - intValue = 0.0 ) Then
			strReturn = strReturn & ".0"
		End If
	End If
	
	'Assign the Return value and Exit
	IeeeHexToStr = strReturn
	Exit Function
End Function