Sub WinStart()
	If InStr(UCase(WScript.FullName), "CSCRIPT.EXE") = 0 Then' If not running under Cscript run the following startup code
		
		
		strRun = "%comspec% /k cscript.exe " & Chr(34) & WScript.ScriptFullName & Chr(34)' Start building the Run script
		
		If objArgs.Count < 5 Then' If the required argument were entered on the command line skip the GUI stuff
			While strServer = ""
				 strMenu = VbCrLf &_
				 "                  Enter the Server Name" 
				strServer = InputBox(strMenu,"Configuration Listing - Server Selection","Server Name")
			Wend
			While strDB = ""
				 strMenu = VbCrLf &_
				 "                Enter the SQL Database Name" 
				strDB = InputBox(strMenu,"Configuration Listing - Database Selection","DB Name")
			Wend
			While strUser = ""
				 strMenu = VbCrLf &_
				 "                 Enter the SQL User Name" 
				strUser = InputBox(strMenu,"Configuration Listing - SQL User Selection","User Name")
			Wend
			While strPwd = ""
				 strMenu = VbCrLf &_
				 "                 Enter the SQL Password" 
				strPwd = InputBox(strMenu,"Configuration Listing - SQL Password","SQL Password")
			Wend
			While strBaseArea = ""
				 strMenu = VbCrLf &_
				 "         Enter the Base Area Hierarchical Name"
				strBaseArea = InputBox(strMenu,"Configuration Listing - Base Area Hier Name","Base Area")
			Wend
			
			strMenu = VbCrLf &_
			"Valid Entries are" & Chr(34) & "NoAttrib" & Chr(34) & " and //x at the end to invoke installed debugger"
			strOptArgs = InputBox(strMenu,"Configuration Listing - Optional Arguments","")
			strRun = strRun & " " & strServer & " " & strDB & " " & strUser & " " & strPwd &_ 
			" " & strBaseArea
			If strOptArgs <> "" Then
				strRun = strRun & " " & strOptArgs
			End If
		Else' If you have the required commandline arguments build the Run script using them
			For x = 0 to (objArgs.Count - 1)
				strRun = strRun & " " & objArgs(x)
			Next
		End If
		
		Set objShell = CreateObject("WScript.Shell") ' Spawn another cmd shell and run the Run script under Cscript
		objShell.Run strRun, 1, False
		WScript.Quit() ' Quit the current Wscript session
	End If
End Sub