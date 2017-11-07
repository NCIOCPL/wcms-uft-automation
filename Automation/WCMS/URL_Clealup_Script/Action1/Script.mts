datatable.Import "L:\OCPL\ODDC\CPSB\QAStuff\New_QAScripts\Automation\WCMS\DateSheets\SummaryMoves\Redirect_Map_Desktop_Test.xls"
intRows = datatable.GetRowCount 
For i = 1 To intRows Step 1
	datatable.SetCurrentRow(i)

	strURL = "http://www.cancer.gov" & datatable.Value("URL")
'	NewURL = datatable.Value("NEW_URL")
	
	If  datatable.Value("Status") <> "Done" Then
		With Browser("micclass:=browser", "index:=0").Page("micclass:=page", "index:=0")	
			Browser("micclass:=browser", "index:=0").Navigate strURL
'			If Browser("micclass:=browser").WinObject("nativeclass:=window","visible:=True","acc_name:=Notification bar").Exist(0) = True  Then
'				Browser("micclass:=browser").WinObject("nativeclass:=window","visible:=True","acc_name:=Notification bar").WinButton("object class:=push button","acc_name:=Cancel").Click
'				strResult = "File download pop-up detected"
'			Else
				BrowserSync
				strResult = CheckForErrorOnPage
'				landingURL = Browser("micclass:=browser", "index:=0").Page("micclass:=page", "index:=0").GetROProperty("url")
'				landingURL = Replace(landingURL,"http://www-new.cancer.gov", "")
'				If landingURL <> NewURL Then
'					URL_Match = "Yes"
'				Else
'					URL_Match = "No"
'				End If
'				msgbox "check the site"
'			End If
		End With

		datatable.Value("Status") = "Done"
'		datatable.Value("Landing_URL") = landingURL
'		datatable.Value("URL_Match") = URL_Match
		datatable.Value("Comments") = strResult

' For updating datasheet after every 15 rows
'		If i Mod 15 = 0  Then
'			datatable.Export "L:\OCPL\ODDC\CPSB\QAStuff\Veena\Automation\Projects\WCMS\Datasheets\SummaryMoves\Redirect_Map_Desktop.xls"	
'		End If
		datatable.Export "L:\OCPL\ODDC\CPSB\QAStuff\New_QAScripts\Automation\WCMS\DateSheets\SummaryMoves\Redirect_Map_Desktop_Test.xls"	
		'Clear Results
		strResult = ""
	End If
Next
