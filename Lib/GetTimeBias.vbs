Sub GetTimeBias()
	Set objShell = CreateObject("WScript.Shell")
	lngDateBiasKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\TimeZoneInformation\Bias")
	If UCase(TypeName(lngDateBiasKey)) = "LONG" Then
		lngDateBias = CLng(lngDateBiasKey)
	ElseIf UCase(TypeName(lngDateBiasKey)) = "VARIANT()" Then
		lngDateBias = CLng(0)
		for x = 0 to UBound(lngDateBiasKey)
			lngDateBias = lngDateBias + CLng((lngDateBiasKey(x) * 256^x ))
		Next
	End If
	
	lngActiveTimeBiasKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias")
	If UCase(TypeName(lngActiveTimeBiasKey)) = "LONG" Then
		lngActiveTimeBias = CLng(lngActiveTimeBiasKey)
	ElseIf UCase(TypeName(lngActiveTimeBiasKey)) = "VARIANT()" Then
		lngActiveTimeBias = CLng(0)
		for x = 0 to UBound(lngActiveTimeBiasKey)
			lngActiveTimeBias = lngDateBias + CLng((lngActiveTimeBiasKey(x) * 256^x ))
		Next
	End If
End Sub