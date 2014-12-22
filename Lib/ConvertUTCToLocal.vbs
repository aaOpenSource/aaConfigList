Function ConvertUTCToLocal(strHexDate, lngBias)
	Dim dblHigh, dblLow, lngHigh, lngLow, i, strCode
	dblHigh = CDbl(0)
	dblLow = CDbl(0)
	
	strCode = "&H" & Mid(strHexDate,1,8)
	dblHigh = CDbl(strCode)
	
	strCode = "&H" & Mid(strHexDate,9,8)
	lngLow = CLng(strCode)
	If lngLow And &h80000000 Then
		dblLow = 2147483648
		lngLow = lngLow And &h7FFFFFFF
		dblLow = dblLow + CDbl(lngLow)
	Else
		dblLow = CDbl(lngLow)
	End If
	
	ConvertUTCToLocal = Cstr(#1/1/1601# + (((dblHigh * (2^32)) + dblLow)/600000000 - lngBias)/1440)
	
End Function
