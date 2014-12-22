	'Returns the Status of individual bits of a value
Function BitStat(Value, Bit)
	If Bit = 31 Then
		If Value And &h80000000 Then
			BitStat = 1
		Else
			BitStat = 0
		End If
	Else
		If Value And (2^Bit) Then
			BitStat = 1
		Else
			BitStat = 0
		End If
	End If
	Exit Function
End Function

