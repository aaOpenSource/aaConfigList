Option Explicit
'***********************************************************************************************************************************	
'*                                                 Archestra Configuration Listing                                                 *
'*	Purpose: This script will list all instances, their UDA's and scripts starting at the level defined in the Base Area. If the   *
'*  script is started with all the required arguments, the script will skip the data entry GUI segment.                            *
'*                                                                                                                                 *
'*                                                                                                                                 *
'***********************************************************************************************************************************

'****************************************************  Start Up Stuff  *************************************************************
	Dim objFileSys
	Set objFileSys = CreateObject("Scripting.FileSystemObject")
	
	Sub Include(strFile)
		Dim objFile, strScript
		If objFileSys.FileExists(strFile) Then
			Set objFile = objFileSys.OpenTextFile(strFile)
			strScript = objFile.ReadAll
			objFile.Close
			ExecuteGlobal strScript
		End If
		Set objFile = Nothing
	End Sub
		
	Include("Lib\Dims.vbs")
	Include("Lib\IeeeHexToStr.vbs")
	Include("Lib\ConvertUTCToLocal.vbs")	
	Include("Lib\BitStat.vbs")
	Include("Lib\ProcessAttributes.vbs")
	Include("Lib\WinStart.vbs")
	Include("Lib\GetTimeBias.vbs")
	
'*****************************************   To debug the subroutines in the lib directory.    **************************************
'* 1. Copy the subroutine to be debugged into this module. 2. Comment out the above include call for that file 
'* 3. run this modulewith the //x argument
'************************************************************************************************************************************

'****************************************************** Program Main Body ***********************************************************
    Set objArgs = WScript.Arguments
	
	strServer = ""
	strDB = ""
	strUser = ""
	strPwd = ""
	strBaseArea = ""
	strUDA = ""
	strOptArgs = ""
	
    WinStart
    GetTimeBias
    
	strBaseArea  = Mid(objArgs(4), (InStrRev(objArgs(4), ".",Len(objArgs(4)))+ 1),(Len(objArgs(4))- InStrRev(objArgs(4), ".",Len(objArgs(4)))))
    strArea = Mid(objArgs(4), (InStrRev(objArgs(4), ".",Len(objArgs(4)))+ 1),(Len(objArgs(4))- InStrRev(objArgs(4), ".",Len(objArgs(4)))))
    strYear = Year(Now())
    strMonth = Month(Now())
    strDay = Day(Now())
    strHour = Mid(FormatDatetime(Now(),4),1,2)
    strMinute = Mid(FormatDatetime(Now(),4),4,2)
    
    If Len(strMonth) < 2 Then
        strMonth = "0" & strMonth
    End If
    
    If Len(strDay) < 2 Then
        strDay = "0" & strDay
    End If
    
	If (objArgs.count > 5) Then
		If InStr(objArgs(5), "NoAttrib") Then
			blnAttrib = False
		Else 
			blnAttrib = True
		End If
	Else
		blnAttrib = True
	End If

    objFileSys.createFolder( "Log\" & strBaseArea & " Config " & strYear & "-" & strMonth & "-" & strDay & " " & strHour & strMinute)
	Set objShell = CreateObject("Wscript.Shell")
	strBaseDir = objShell.CurrentDirectory & "\Log\" & strBaseArea & " Config " & strYear & "-" & strMonth & "-" & strDay & " " & strHour & strMinute
	objShell.CurrentDirectory = strBaseDir
    Set objADOConn = CreateObject("ADODB.Connection")
	dblValue = CDbl(0)
	lngObjCnt = CLng(0)
	lngAttrCnt = CLng(0)
	
	datStart = Now()
			
	Wscript.StdOut.Write "                             Archestra Configuration Listing" & Chr(13)
	
	Wscript.StdOut.Write VbCrLf
	Wscript.StdOut.Write VbCrLf
	
	Wscript.StdOut.Write "    Server    : " & objArgs(0) & Chr(13)
	Wscript.StdOut.Write VbCrLf
	Wscript.StdOut.Write "    Database  : " & objArgs(1) & Chr(13)
	Wscript.StdOut.Write VbCrLf
	Wscript.StdOut.Write "    Base Area : " & objArgs(4) & Chr(13)
	
	Wscript.StdOut.Write VbCrLf
	Wscript.StdOut.Write VbCrLf
	
	strConnectionString = "driver={SQL Server};server=" & objArgs(0) & ";network=dbmssocn;uid=" & objArgs(2) & ";pwd=" & objArgs(3) & ";"
       
    objADOConn.Open strConnectionString ' Connect to SQL Server
	
	strSQL = "use " & objArgs(1) ' Set the Database to be used

    objADOConn.execute(strSQL)
	
	strSQL = "SET NOCOUNT ON"

    objADOConn.execute(strSQL)
	
    strSQL = "SELECT  go.hierarchical_name, go.gobject_id,  go.derived_from_gobject_id, go.template_definition_id, td.category_id, go.checked_in_package_id, go.deployed_package_id, go.configuration_version, go.deployed_version FROM "& objArgs(1) &".dbo.gobject go Inner Join "& objArgs(1) &".dbo.Template_Definition td on go.template_definition_id = td.template_definition_id WHERE(hierarchical_name LIKE '" & objArgs(4) & "%') ORDER BY go.hierarchical_name"

    Set objADOrs0 = objADOConn.execute(strSQL)
	
	Do While Not objADOrs0.EOF ' Loop through all records of listed attributes
		intGobject_Id = objADOrs0("gobject_id")
		intDFGobject_Id = objADOrs0("derived_from_gobject_id")
		intTmpltDef_Id = objADOrs0("template_definition_id")
		intCtgry_Id = objADOrs0("category_id")
		intConfig_Ver = objADOrs0("configuration_version")
		intDeploy_Ver = objADOrs0("deployed_version")
		intChkd_Pack_Id = objADOrs0("checked_in_package_id")
		intDep_Pack_Id = objADOrs0("deployed_package_id")
		strHier_Name = objADOrs0("hierarchical_name")
		'First Pass
		strText = "   " & objArgs(1) & "." & strHier_Name
		lngObjCnt = lngObjCnt + 1
		x = 1
		strArea = ""
		Do While y < Len(Trim(strText))
			x = InStr(x, Trim(strText), ".")
			y = InStr((x + 1), Trim(strText), ".")
			If x > 0 Then
				If y > 0 Then
					strArea = strArea & Mid(Trim(strText), (x + 1),(y - (x + 1))) & "-"
				Else
					strArea = strArea & Mid(Trim(strText), (x + 1))
					Exit Do
				End If
			Else
				Exit Do
			End If
			x = x + 1
		Loop
		If Not objFileSys.FileExists( strArea & ".txt") Then
			If IsObject("objFile") Then
				objFile.Close
			End If 
			Set objFile = objFileSys.createtextfile( strArea & ".txt",True)
		End If
		objFile.WriteLine (Chr(13))
		If (objADOrs0("deployed_package_id") > 0) Then
			objFile.WriteLine ("   <------------------->  Object_ID = " & intGobject_Id & " - Deployed_Version = " & intDeploy_Ver & " <------------------->" & Chr(13))
		Else
			objFile.WriteLine ("   <------------------->  Object_ID = " & intGobject_Id & " - Configuration_Version = " & intConfig_Ver & " <------------------->" & Chr(13))
		End If
		objFile.WriteLine (strText & Chr(13)) ' Write to Log
		objFile.WriteLine (Chr(13))
		If blnAttrib Then
			ProcessAttributes
		End If
		
		strSQL = "SELECT  go.tag_name, go.gobject_id, go.derived_from_gobject_id, go.template_definition_id, td.category_id, go.checked_in_package_id, go.deployed_package_id, go.configuration_version, go.deployed_version FROM "& objArgs(1) &".dbo.gobject go Inner Join "& objArgs(1) &".dbo.Template_Definition td on go.template_definition_id = td.template_definition_id WHERE((go.area_gobject_id = " & objADOrs0("gobject_id") & ") AND (go.contained_by_gobject_id = 0) AND (go.hierarchical_name NOT LIKE '" & objArgs(4) & "')) ORDER BY go.tag_name"
		Set objADOrs2 = objADOConn.execute(strSQL)
		intLevel = 0
		x = 0
		Do While Not objADOrs2.EOF
			x = x + 1
			objADOrs2.moveNext
		Loop
		If x > 0 Then
			objADOrs2.MoveFirst
			Redim intGobject_Id_0(x-1)
			Redim intDFGobject_Id_0(x-1)
			Redim intTmpltDef_Id_0(x-1)
			Redim intCtgry_Id_0(x-1)
			Redim intDeploy_Ver_0(x-1)
			Redim intConfig_Ver_0(x-1)
			Redim intChkd_Pack_id_0(x-1)
			Redim intDep_Pack_Id_0(x-1)
			Redim strTag_Name_0(x-1)
			Redim intPointer(0)
			intPointer(0) = 0
			x = 0
			Do While Not objADOrs2.EOF
				intGobject_id_0(x) = objADOrs2("gobject_id")
				intDFGobject_id_0(x) = objADOrs2("derived_from_gobject_id")
				intTmpltDef_id_0(x) = objADOrs2("template_definition_id")
				intCtgry_id_0(x) = objADOrs2("category_id")
				intDeploy_Ver_0(x) = objADOrs2("deployed_version")
				intConfig_Ver_0(x) = objADOrs2("configuration_version")
				intChkd_Pack_id_0(x) = objADOrs2("checked_in_package_id")
				intDep_Pack_id_0(x) = objADOrs2("deployed_package_id")
				strTag_Name_0(x) = objADOrs2("tag_name")
				x = x + 1
				objADOrs2.moveNext
			Loop
			Do While intPointer(0) <= UBound(intGobject_id_0)
				Execute("intGobject_Id = intGobject_id_" & intLevel & "(" & CStr(intPointer(intLevel)) & ")")
				Execute("intDFGobject_Id = intDFGobject_id_" & intLevel & "(" & CStr(intPointer(intLevel)) & ")")
				Execute("intTmpltDef_Id = intTmpltDef_id_" & intLevel & "(" & CStr(intPointer(intLevel)) & ")")
				Execute("intCtgry_Id = intCtgry_id_" & intLevel & "(" & CStr(intPointer(intLevel)) & ")")
				Execute("intConfig_Ver = intConfig_Ver_" & intLevel & "(" & CStr(intPointer(intLevel)) & ")")
				Execute("intDeploy_Ver = intDeploy_Ver_" & intLevel & "(" & CStr(intPointer(intLevel)) & ")")
				Execute("intChkd_Pack_Id = intChkd_Pack_id_" & intLevel & "(" & CStr(intPointer(intLevel)) & ")")
				Execute("intDep_Pack_Id = intDep_Pack_id_" & intLevel & "(" & CStr(intPointer(intLevel)) & ")")
				strHier_Name = objADOrs0("hierarchical_name")
				For x = 0 to intLevel
					strScratch = Eval("strTag_Name_" & CStr(x) & "(" & CStr(intPointer(x)) & ")" )
					strHier_Name = strHier_Name & "." & strScratch
				Next
				strText = "   " & objArgs(1) & "." & strHier_Name
				lngObjCnt = lngObjCnt + 1
				'Wscript.StdOut.WriteLine Chr(13) & strText ' Write to the Console
				'Wscript.StdOut.WriteLine ""
				objFile.WriteLine (Chr(13))
				If (intDep_Pack_Id > 0) Then
					objFile.WriteLine ("   <------------------->  Object_ID = " & CStr(intGobject_Id) &  " - Deployed_Version = " & CStr(intDeploy_Ver) & " <------------------->" & Chr(13))
				Else
					objFile.WriteLine ("   <------------------->  Object_ID = " & CStr(intGobject_Id) &  " - Configuration_Version = " & CStr(intConfig_Ver) & " <------------------->" & Chr(13))
				End If
				objFile.WriteLine (strText & Chr(13)) ' Write to Log
				objFile.WriteLine (Chr(13))
				If blnAttrib Then
					ProcessAttributes
				End If				
				strSQL = "SELECT  go.tag_name, go.gobject_id, go.derived_from_gobject_id, go.template_definition_id, td.category_id, go.checked_in_package_id, go.deployed_package_id, go.configuration_version, go.deployed_version FROM "& objArgs(1) &".dbo.gobject go Inner Join "& objArgs(1) &".dbo.Template_Definition td on go.template_definition_id = td.template_definition_id WHERE((go.area_gobject_id = " & objADOrs0("gobject_id") & ") AND (go.contained_by_gobject_id = " & intGobject_Id & ") AND (go.hierarchical_name NOT LIKE '" & objArgs(4) & "')) ORDER BY go.tag_name"
				Set objADOrs2 = objADOConn.execute(strSQL)
				x = 0
				Do While Not objADOrs2.EOF
					x = x + 1
					objADOrs2.moveNext
				Loop
				If x > 0 Then
					objADOrs2.MoveFirst
					intLevel = intLevel + 1
					Execute("Redim intGobject_id_" & CStr(intLevel) & "(" & CStr(x - 1) & ")" )
					Execute("Redim intDFGobject_Id_" & CStr(intLevel) & "(" & CStr(x - 1) & ")" )
					Execute("Redim intTmpltDef_Id_" & CStr(intLevel) & "(" & CStr(x - 1) & ")" )
					Execute("Redim intCtgry_id_" & CStr(intLevel) & "(" & CStr(x - 1) & ")" )
					Execute("Redim intConfig_Ver_" & CStr(intLevel) & "(" & CStr(x - 1) & ")" )
					Execute("Redim intDeploy_Ver_" & CStr(intLevel) & "(" & CStr(x - 1) & ")" )
					Execute("Redim intChkd_Pack_id_" & CStr(intLevel) & "(" & CStr(x - 1) & ")" )
					Execute("Redim intDep_Pack_id_" & CStr(intLevel) & "(" & CStr(x - 1) & ")" )
					Execute("Redim strTag_Name_" & CStr(intLevel) & "(" & CStr(x - 1) & ")" )
					Redim Preserve intPointer(intLevel)
					intPointer(intLevel) = 0
					x = 0
					Do While Not objADOrs2.EOF
						Execute ("intGobject_id_" & CStr(intLevel) & "(" & CStr(x) & ") = " & CStr(objADOrs2("gobject_id")))
						Execute ("intDFGobject_Id_" & CStr(intLevel) & "(" & CStr(x) & ") = " & CStr(objADOrs2("derived_from_gobject_id")))
						Execute ("intTmpltDef_Id_" & CStr(intLevel) & "(" & CStr(x) & ") = " & CStr(objADOrs2("template_definition_id")))
						Execute ("intCtgry_id_" & CStr(intLevel) & "(" & CStr(x) & ") = " & CStr(objADOrs2("category_id")))
						Execute ("intConfig_Ver_" & CStr(intLevel) & "(" & CStr(x) & ") = " & Chr(34) & objADOrs2("configuration_version") & Chr(34))
						Execute ("intDeploy_Ver_" & CStr(intLevel) & "(" & CStr(x) & ") = " & Chr(34) & objADOrs2("deployed_version") & Chr(34))
						Execute ("intChkd_Pack_id_" & CStr(intLevel) & "(" & CStr(x) & ") = " & Chr(34) & objADOrs2("checked_in_package_id") & Chr(34))
						Execute ("intDep_Pack_id_" & CStr(intLevel) & "(" & CStr(x) & ") = " & Chr(34) & objADOrs2("deployed_package_id") & Chr(34))
						Execute ("strTag_Name_" & CStr(intLevel) & "(" & CStr(x) & ") = " & Chr(34) & objADOrs2("tag_name") & Chr(34))
						x = x + 1
						objADOrs2.moveNext
					Loop
				Else
					intBound = Eval("UBound(intGobject_id_" & CStr(intLevel) & ")")
					Do While intPointer(intLevel) = intBound
						If intLevel > 0 Then
							intLevel = intLevel - 1
							Redim Preserve intPointer(intLevel)
							intBound = Eval("UBound(intGobject_id_" & CStr(intLevel) & ")")
						Else
							Exit Do
						End If
					Loop
					If (intLevel = 0) Then
						intPointer(intLevel) = intPointer(intLevel) + 1
					Else
						If (intPointer(intLevel) < intBound) Then
							intPointer(intLevel) = intPointer(intLevel) + 1
						End If
					End If
				End If
			Loop
		End If
		objADOrs0.moveNext
	Loop
	
	Wscript.StdOut.Write VbCrLf
	Wscript.StdOut.Write VbCrLf

	objFile.Close
	Set objFile = Nothing
