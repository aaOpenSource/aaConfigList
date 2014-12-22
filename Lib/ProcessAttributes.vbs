Sub ProcessAttributes()
	Dim blnRunOnce1, blnRunOnce2, z, strPrimDefName, strPrimName, strIntVal, blnSTX, blnETX
	
	If (objADOrs0("deployed_package_id") > 0) Then
		intPack_ID = intDep_Pack_Id
	Else
		intPack_ID = intChkd_Pack_Id
	End If
	
	strSQL = "Select da.gobject_id, da.package_id, da.mx_primitive_id, da.attribute_name, pi.primitive_name, pd.primitive_name primitive_def_name, da.mx_data_type, da.is_array, da.mx_value From " &_
		objArgs(1) & ".dbo.Dynamic_Attribute da Inner Join "& objArgs(1) &".dbo.primitive_instance pi on  da.mx_primitive_id = pi.mx_primitive_id and da.gobject_id = pi.gobject_id Inner Join "&_
		objArgs(1) & ".dbo.primitive_definition pd on  pi.primitive_definition_id = pd.primitive_definition_id where da.gobject_id = " & intGobject_Id & " and  da.package_id = "& intPack_ID &_
		" and da.mx_value Not Like '0x0506000000020000000000' and da.mx_value Not Like '0x0534000000300000003C004900740065006D0073004C006900730074003E003C002F004900740065006D0073004C006900730074003E000000'" &_
		" Order By primitive_def_name, primitive_name, attribute_name "
	Set objADOrs1 = objADOConn.execute(strSQL)
	
	If (Not (objADOrs1.BOF And objADOrs1.EOF)) Then
		strSQL = "SELECT  extension_type, primitive_name FROM "& objArgs(1) &".dbo.primitive_instance WHERE((gobject_id = " & objADOrs1("gobject_id") & ") And (package_id = " & objADOrs1("package_id") & ") And (extension_type <> '')) Order By extension_type"
		Set objADOrs3 = objADOConn.execute(strSQL)
	End If
'	For y = 0 to 1
'		Select Case y
'			Case 0
				strSQL = "Select ta.gobject_id, ta.mx_primitive_id, pi.primitive_name, pd.primitive_name primitive_def_name, ta.mx_data_type, ta.mx_value From "& objArgs(1) &".dbo.Template_Attribute ta Inner Join "& objArgs(1) &".dbo.primitive_instance pi on  ta.mx_primitive_id = pi.mx_primitive_id and ta.gobject_id = pi.gobject_id Inner Join "& objArgs(1) &".dbo.primitive_definition pd on  pi.primitive_definition_id = pd.primitive_definition_id where ta.gobject_id = " & intDFGobject_Id &" and ta.mx_data_type = 16 and  ta.mx_value Not Like '0x0506000000020000000000' and ta.mx_value Not Like '0x0534000000300000003C004900740065006D0073004C006900730074003E003C002F004900740065006D0073004C006900730074003E000000'"
'			Case 1
'				strSQL = "Select gobject_id, mx_primitive_id, primitive_name, extension_type primitive_def_name, primitive_attributes mx_value From "& objArgs(1) &".dbo.primitive_instance where gobject_id = " & intGobject_Id & " and created_by_parent = 0 And Package_id = " & intPack_ID
'			Case Else
'		End Select
		Set objADOrs4 = objADOConn.execute(strSQL)
		
		
		If Not objADOrs4.EOF Then
			Set objDict1 = CreateObject("Scripting.Dictionary")
			Set objDict2 = CreateObject("Scripting.Dictionary")
		End If
		
		Do While Not objADOrs4.EOF ' Loop through all records of listed attributes
'			If y = 1 Then
'				Set objADOStream = CreateObject("ADODB.Stream")
'				objADOrs4.moveFirst
'				objADOrs4.moveNext
'				objADOStream.Open
'				objADOStream.Type = 1 '1-Binary, 2-Text
'				objADOStream.Write objADOrs4("mx_value")
'				objADOStream.Position = 0
'				objADOStream.Type = 2'1-Binary, 2-Text
'				objADOStream.CharSet = "us-ascii"
'				blnSTX = False
'				blnETX = False
'				Do While objADOStream.Position < objADOStream.Size
'					strValue = objADOStream.ReadText(1)
'					If (blnSTX = True) And (blnETX = False) Then
'						If ((asc(strValue) > 31) And (asc(strValue) < 127)) Or (asc(strValue) = 9) Then
'							strRawData = strRawData & strValue
'							If InStr(strRawData,"_ScriptExtension") Then
'								strRawData = Mid(strRawData, 3, (InStr(strRawData,"_ScriptExtension")-3))
'								objFile.WriteLine ("    " & Chr(13))
'								objFile.WriteLine ("                                    **** " & strRawData & " ****" & Chr(13)) 
'								objFile.WriteLine ("    " & Chr(13))
'								strRawData = ""
'							End If
'						Else
'							If	(asc(strValue) = 10) Or (asc(strValue) = 3) Then
'								objFile.WriteLine ("    " & strRawData & Chr(13))
'								strRawData = ""
'								If (asc(strValue) = 3) Then
'									blnETX = True
'									objFile.WriteLine ("    " & Chr(13))
'								End If
'							End If
'						End If
'					Else
'						If (asc(strValue) = 2) Then
'							blnSTX = True
'							blnETX = False
'							strRawData = ""
'						End If
'					End If
'				Loop
'				objADOStream.Close
'				Set objADOStream = Nothing
'			Else
				strRawData = objADOrs4("mx_value")
'			End If
			If Not (IsNull(strRawData)) Then
				If (Len(strRawData) > 0) Then
					strRawData = Mid(strRawData,19, (Len(strRawData) - 24))
					For x = 1 to Len(strRawData) Step 4
						If (Len(strRawData) - x) > 3 Then
							If objADOrs4("primitive_def_name") = "ScriptExtension" Then
								If strPrimDefName <> "ScriptExtension" Then
									objFile.WriteLine ("    " & Chr(13))
									objFile.WriteLine ("                              <---------- Scripts ---------->" & Chr(13))
									objFile.WriteLine ("    " & Chr(13))
									strPrimDefName = "ScriptExtension"
								End If
								If strPrimName <> objADOrs4("primitive_name") Then
									objFile.WriteLine ("    " & Chr(13))
									objFile.WriteLine ("                                    **** " & objADOrs4("primitive_name") & " ****" & Chr(13)) 
									objFile.WriteLine ("    " & Chr(13))
									strPrimName = objADOrs4("primitive_name")
								End If
								If (Mid(strRawData,(x + 2),2) = "0A") Then
									strIntVal = strIntVal & Mid(strRawData,(x + 2),2)
									strValue =  strValue & Chr(13)
									objFile.WriteLine ("    " & strValue)
									strValue = ""
								Else
									StrCode = "Chr(&h" & Mid(strRawData,(x + 2),2) & ")"
									strIntVal = strIntVal & Mid(strRawData,(x + 2),2)
									strValue =  strValue & Eval(strCode)
								End If
							Else
								StrCode = "Chr(&h" & Mid(strRawData,(x + 2),2) & ")"
								strIntVal = strIntVal & Mid(strRawData,(x + 2),2)
								strValue =  strValue & Eval(strCode)
							End If
							If objADOrs4("primitive_def_name") = "SG" Then
								If (InStr(strValue, "<Item Name=") > 0) And (InStr(strValue, "/>") > 0) Then
									strItem = Mid(strValue,(InStr(strValue, "Item Name=")+11),(InStr((InStr(strValue, "Item Name=")+12),strValue, " Alias=" )-(InStr(strValue, "Item Name=")+12)))
									strAlias = Mid(strValue,(InStr(strValue, " Alias=")+8),(InStr((InStr(strValue, " Alias=")+9),strValue, "/" ) - (InStr(strValue, " Alias=")+9)))
									If objDict1.Exists(Chr(34) & strAlias & Chr(34)) Then
										objDict1.Remove(Chr(34) & strAlias & Chr(34))
									End If
									objDict1.Add (Chr(34) & strAlias & Chr(34)), (Chr(34) & strItem & Chr(34))
									If objDict2.Exists(Chr(34) & strAlias & Chr(34)) Then
										objDict2.Remove(Chr(34) & strAlias & Chr(34))
									End If
									objDict2.Add (Chr(34) & strAlias & Chr(34)), (Chr(34) & objADOrs4("primitive_name") & Chr(34))
									strValue = ""
								End If
							End If
						End If
					Next
					If objADOrs4("primitive_def_name") = "ScriptExtension" Then
						lngAttrCnt = lngAttrCnt + 1
					End If
					strIntVal = Empty
					strValue = Empty
					objADOrs4.moveNext
				End If
			End If
		Loop
	'Next 'y
	blnRunOnce1 = True
	Do While Not objADOrs1.EOF ' Loop through all records of listed attributes
		If blnRunOnce1 Then
			objFile.WriteLine (Chr(13))
			objFile.WriteLine ("                            <---------- Attributes ---------->" & Chr(13))
			objFile.WriteLine (Chr(13))
			blnRunOnce1 = False
		End If
		strRawData = objADOrs1("mx_value") ' Get the Attribute data
		 strValue = ""
		Select Case objADOrs1("mx_data_type")
			Case 1
				If objADOrs1("is_array") Then
					StrCode = "CLng(&h" & Mid(strRawData, 11, 4) & ")"
					intSize = Eval(strCode)
					strValue = ""
					intLocation = 21
					For x = 1 to (intSize)
						StrCode = "CLng(&h" & Mid(strRawData, (intLocation + (x * 4)), 4) & ")"
						intValue = Eval(strCode)
						If (intValue < 0) Then
							strValue = strValue & " " & CStr(x) & ": True"
						Else
							strValue = strValue & " " & CStr(x) & ": False"
						End If
						if (x < intSize) Then
							strValue = strValue & ","
						End If
					Next
				Else
					StrCode = "CInt(&h" & Mid(strRawData, 5, 2) & ")"
					intValue = Eval(strCode)
					If (intValue > 0) Then
						strValue = "True"
					Else
						strValue = "False"
					End If
				End If
			Case 2
				If objADOrs1("is_array") Then
					StrCode = "CLng(&h" & Mid(strRawData, 11, 4) & ")"
					intSize = Eval(strCode)
					strValue = ""
					intLocation = 23
					For x = 1 to (intSize)
						StrCode = "CLng(&h" & Mid(strRawData, (intLocation + (x * 8)), 2) & Mid(strRawData, (intLocation + (x * 8) -2), 2) & Mid(strRawData, (intLocation + (x * 8) -4), 2) & Mid(strRawData, (intLocation + (x * 8) -6), 2) & ")"
						intValue = Eval(strCode)
						strValue = strValue & " " & CStr(x) & ": " & CStr(intValue)
						if (x < intSize) Then
							strValue = strValue & ","
						End If
					Next
				Else
					StrCode = "CLng(&h" & Mid(strRawData, 11, 2) & Mid(strRawData, 9, 2) & Mid(strRawData, 7, 2) & Mid(strRawData, 5, 2) & ")"
					intValue = Eval(strCode)
					strValue = Cstr(intValue)
				End If
			Case 3
				If objADOrs1("is_array") Then
					StrCode = "CLng(&h" & Mid(strRawData, 11, 4) & ")"
					intSize = Eval(strCode)
					strValue = ""
					intLocation = 25
					For x = 1 to (intSize)
						strCode = Mid(strRawData, (intLocation + 6), 2) & Mid(strRawData, (intLocation + 4), 2) & Mid(strRawData, (intLocation + 2), 2) & Mid(strRawData, intLocation, 2)
						strValue = strValue & " " & CStr(x) & ": " & IeeeHexToStr(strCode)
						if (x < intSize) Then
							strValue = strValue & ","
						End If
						intLocation = intLocation + 8
					Next
				Else
					strCode = Mid(strRawData, 11, 2) & Mid(strRawData, 9, 2) & Mid(strRawData, 7, 2) & Mid(strRawData, 5, 2)
					strValue = IeeeHexToStr(strCode)
				End If
			Case 4
				If objADOrs1("is_array") Then
					StrCode = "CLng(&h" & Mid(strRawData, 11, 4) & ")"
					intSize = Eval(strCode)
					strValue = ""
					intLocation = 25
					For x = 1 to (intSize)
						strCode = Mid(strRawData, (intLocation + 14), 2) & Mid(strRawData, (intLocation + 12), 2) & Mid(strRawData, (intLocation + 10), 2) & Mid(strRawData, (intLocation + 8), 2) &_
						Mid(strRawData, (intLocation + 6), 2) & Mid(strRawData, (intLocation + 4), 2) & Mid(strRawData, (intLocation + 2), 2) & Mid(strRawData, intLocation, 2)
						strValue = strValue & " " & CStr(x) & ": " & IeeeHexToStr(strCode)
						if (x < intSize) Then
							strValue = strValue & ","
						End If
						intLocation = intLocation + 16
					Next
				Else
					strCode = Mid(strRawData, 19, 2) & Mid(strRawData, 17, 2) & Mid(strRawData, 15, 2) &_
					Mid(strRawData, 13, 2) & Mid(strRawData, 11, 2) & Mid(strRawData, 9, 2) & Mid(strRawData, 7, 2) &_
					Mid(strRawData, 5, 2)
					strValue = IeeeHexToStr(strCode)
				End If
			Case 5	
				If objADOrs1("is_array") Then
					strEmpty = "00B00000005060000000200000000"
					StrCode = "CInt(&h" & Mid(strRawData, 11, 4) & ")"
					intSize = Eval(strCode)
					Redim intLocations(intSize - 1)
					intLocation = 23
					For x = 0 to (intSize - 1)
						intLocation = (InStr(intLocation, strRawData, "0000000") - 3)
						If (InStr(Mid(strRawData, (intLocation + 10), 15),"0000000") And InStr(Mid(strRawData, (intLocation + 29), 6),"0000"))Then
							intLocation = (InStr((intLocation + 10),strRawData, "0000000") - 3)
						End If
						strTest = Mid(strRawData, (intLocation + 8), 29)
						If InStr(Mid(strRawData, intLocation + 8, 35), strEmpty) Then
							intLocation = InStr(intLocation + 8, strRawData, strEmpty)
							intLocations(x) = intLocation * -1
							intLocation = intLocation + 25
						Else
							intLocations(x) = intLocation
							intLocation = intLocation + 25
						End If
					Next
					strValue = ""
					For x = 0 to Ubound(intLocations)
						If intLocations(x) > 0 then
							strValue =  strValue & " " & CStr(x + 1) & ":" & Chr(34)
							While Mid(strRawData, (intLocations(x) + 25), 4) <> "0000"
								StrCode = "Chr(&h" & Mid(strRawData, (intLocations(x) + 25), 4) & ")"
								strValue =  strValue & Eval(strCode) 
								intLocations(x) = intLocations(x) + 4
							WEnd
							If x < Ubound(intLocations) Then
								strValue = strValue & Chr(34) & ","
							Else
								strValue = strValue & Chr(34)
							End If
						Else
							If x < Ubound(intLocations) Then
								strValue =  strValue & " " & CStr(x + 1) & ":" & Chr(34) & Chr(34) & ","
							Else
								strValue =  strValue & " " & CStr(x + 1) & ":" & Chr(34) & Chr(34)
							End If
						End If
					Next
				Else
					If strRawData = "0x0506000000020000000000" Then ' Check for an empty value
						strValue =  Chr(34) & Chr(34)
					Else ' Translate the value from unicode to text
						strRawData = Mid(strRawData,19, (Len(strRawData) - 24))
						For x = 1 to Len(strRawData) Step 4
							StrCode = "Chr(&h" & Mid(strRawData,3,2) & ")"
							strValue =  strValue & Eval(strCode)
							strRawData = Mid(strRawData,5, (Len(strRawData) - 4))
						Next
						strValue = Chr(34) & strValue & Chr(34)
					End If
				End If
			Case 6
				If objADOrs1("is_array") Then
					StrCode = "CLng(&h" & Mid(strRawData, 11, 4) & ")"
					intSize = Eval(strCode)
					strValue = ""
					For x = 0 to (intSize -1)
						intScratch = 24 * x
						strCode = Mid(strRawData, (39) + intScratch, 2) & Mid(strRawData, (37) + intScratch, 2) & Mid(strRawData, (35) + intScratch, 2) &_
									Mid(strRawData, (33) + intScratch, 2) & Mid(strRawData, (31) + intScratch, 2) & Mid(strRawData, (29) + intScratch, 2) &_
									Mid(strRawData, (27) + intScratch, 2) & Mid(strRawData, (25) + intScratch, 2)
									
						strValue = strValue + CStr(x+1) + ": " + ConvertUTCToLocal(strCode, lngActiveTimeBias)
						
						If (x < intSize -1) Then
							strValue = strValue + ", "
						End If
						
						Next
				Else
					strCode = Mid(strRawData, 27, 2) & Mid(strRawData, 25, 2) & Mid(strRawData, 23, 2) & Mid(strRawData, 21, 2) &_
								Mid(strRawData, 19, 2) & Mid(strRawData, 17, 2) & Mid(strRawData, 15, 2) & Mid(strRawData, 13, 2)
								
					strValue = ConvertUTCToLocal(strCode, lngDateBias)
					
				End If
			Case 16
				If (strRawData = "0x0506000000020000000000") or (strRawData = "0x00") Then ' Check for an empty value
					strValue =  Chr(34) & Chr(34)
				Else ' Translate the value from unicode to text
					strRawData = Mid(strRawData,19, (Len(strRawData) - 24))
					For x = 1 to Len(strRawData) Step 4
						If Len(strRawData) >= 4 Then
							StrCode = "Chr(&h" & Mid(strRawData,3,2) & ")"
							strValue =  strValue & Eval(strCode)
							strRawData = Mid(strRawData,5, (Len(strRawData) - 4))
						End If
					Next
				End If			
			Case Else
				StrCode = "CInt(&h" & Mid(strRawData, 10, 4) & ")"
		End Select

		If (intCtgry_Id = 12) and (IsObject(objDict1)) and (IsObject(objDict2)) Then
			If objDict1.Exists(Chr(34) & objADOrs1("attribute_name") & Chr(34)) And objDict1.Exists(Chr(34) & objADOrs1("attribute_name") & Chr(34)) Then
				strText = "   " & mid(objDict2.Item(Chr(34) & objADOrs1("attribute_name") & Chr(34)),2,(Len(objDict2.Item(Chr(34) & objADOrs1("attribute_name") & Chr(34))) - 2)) & "." & objADOrs1("attribute_name") & " = " & mid(objDict1.Item(Chr(34) & objADOrs1("attribute_name") & Chr(34)),2,(Len(objDict1.Item(Chr(34) & objADOrs1("attribute_name") & Chr(34)))-2))
			Else
				strText = "   " & objADOrs1("attribute_name")
				If (objADOrs1("mx_data_type") > 0) Then
					strText = strText & " = " &  strValue
				End If
			End If
		Else
			strText = "   " & objADOrs1("attribute_name")
			If (objADOrs1("mx_data_type") > 0) Then
				strText = strText & " = " &  strValue
			End If
		End If
		
		blnRunOnce2 = True
		If ((objADOrs3.EOF And (Not objADOrs3.BOF)) Or (Not (objADOrs3.BOF And objADOrs3.EOF))) Then
			objADOrs3.MoveFirst
			Do While Not objADOrs3.EOF
				If objADOrs3("primitive_name") = objADOrs1("attribute_name") Then
					If blnRunOnce2 Then
						strText = strText & "          Extension Flags - "
						blnRunOnce2 = False
					Else
						strText = strText & ", "
					End If
					Select Case objADOrs3("extension_type")
						Case "alarmextension"
							strText = strText & "A"
						Case "historyextension"
							strText = strText & "H"
						Case "inputextension"
							strText = strText & "I"
						Case "outputextension"
							strText = strText & "O"
						Case "inputoutputextension"
							strText = strText & "I/O"
						Case "ScriptExtension"
							strText = strText & "S"
						Case Else
							strText = strText & objADOrs3("extension_type")
					End Select
					strText = Replace(strText, "I/O, I/O", "I/O")
				End If
				objADOrs3.moveNext
			Loop
		End If
		lngAttrCnt = lngAttrCnt + 1
		objFile.WriteLine (strText & Chr(13)) ' Write to Log
		For z = 1 to 60000
		Next
		datEnd = Now
		If Not blnNightCheck Then
		    Wscript.StdOut.Write Chr(13) & "         " & CStr(lngObjCnt) & " Objects - "  & CStr(lngAttrCnt) & " Attributes - " & Cstr(Fix(DateDiff("S", datStart, datEnd)/60)) & " minutes : " &_
		    Cstr(DateDiff("S", datStart, datEnd) Mod 60) & " seconds   "
		End If
		objADOrs1.moveNext
	Loop
	Set objDict1 = Nothing
	Set objDict2 = Nothing
End Sub
