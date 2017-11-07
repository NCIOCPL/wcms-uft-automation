 

	
''============================================================
'' Veriable Declaration
''============================================================
Set ParentBrw = Browser("name:=Rhythmyx - Content Explorer").Page("title:=Rhythmyx - Content Explorer").frame("name:=navcontent").JavaApplet("tagname:=PSContentExplorerApplet","to_class:=JavaApplet")
set brMenuList = ParentBrw.JavaMenu("label:=New Item","to_class:=JavaMenu","path:=JMenu;JPopupMenu;JPanel;JLayeredPane;JRootPane;Popup\$HeavyWeightWindow;PluginEmbeddedFrame;")
Set LeftPlanel = ParentBrw.JavaTree("tagname:=Content Path:  ","to_class:=JavaTree")
Set RightPlanel = ParentBrw.Javatable("tagname:=//Sites/.*","to_class:=JavaTable","columns_names:=Content Title;.*Content Type.*","toolkit class:=com\.percussion\.cx\.PSMainDisplayPanel\$1")
Set NewFolderWin =ParentBrw.JavaDialog("tagname:=Create folder","to_class:=JavaDialog")
	CommuTab= "Browser(""title:=Rhythmyx - Content Explorer"").Page(""title:=Rhythmyx - Content Explorer"").frame(""name:=banner"")."
	SelectCommuPage = "Browser(""title:=Community and Language Selection Home Page"").Page(""title:=Community and Language Selection Home Page"")."


	
''============================================================
'' Function Name: PerucssionLogin
''============================================================

Function PercussionLogin(sUsername,sPassword)
		 Browser("micclass:=browser").Page("micclass:=page").WebEdit("html id:=j_username,name:=j_username").set sUsername
		 Browser("micclass:=browser").Page("micclass:=page").WebEdit("name:=j_password").set sPassword
		 Browser("micclass:=browser").Page("micclass:=page").Image("name:=Logon").set Click
End Function

Function CreateFolderContent(FolderName,NewFolderName, Status)
		If  SiteName = FolderName Then
			strPath = NewFolderName
			OpenContentMenu(FolderName)
		Else
			strPath = FolderName & ";" & NewFolderName
			OpenContentMenu(FolderName)
		End If
		
'		Browser("name:=Rhythmyx - Content Explorer").Refresh
		BoolFolderExists = ParentBrw.JavaTree("class description:=list","tagname:=Content Path.*","value:=Content Explorer;Sites;"& SiteName &";" &strPath).Exist(5)
	If BoolFolderExists = False Then
		SelectContentFolder(FolderName)
		Set NewFolderWin =ParentBrw.JavaDialog("tagname:=Create folder","to_class:=JavaDialog")
		Dim strTemp
		Call OpenContentMenu(FolderName)
		wait 1
		
		ParentBrw.JavaMenu("label:=New Folder\.\.\.","to_class:=JavaMenu","path:=JMenuItem;JPopupMenu;JPanel;JLayeredPane;JRootPane;Popup\$HeavyWeightWindow;PluginEmbeddedFrame;").Select
		wait 2
		NewFolderWin.JavaEdit("tagname:=Folder Name:","to_class:=JavaEdit","index:=0").Set NewFolderName
		NewFolderWin.JavaButton("tagname:=OK","to_class:=JavaButton","index:=0").Click
		wait 8
		SelectContentFolder(strPath)
		Wait 4
		If Status <> "Draft" Then
			PushWorkFlowFromDraftToPublic NewFolderName, "Regular"
		End If
	End If

End Function		


''============================================================
'' Function Name: SelectContentFolder
''============================================================
Function SelectContentFolder(FolderName)
		If  SiteName = FolderName Then
			LeftPlanel.Select "Content Explorer;Sites;"& SiteName &""
		Else
			LeftPlanel.Select "Content Explorer;Sites;"& SiteName &";"& FolderName &"" 
'			LeftPlanel.Select "Content Explorer;Sites;" & FolderName &"" 
		End If
		
	''''	SiteName
End Function


'Function SelectLeftPanelMenu(MenuName,More)
''		set brMenuList = ParentBrw.JavaMenu("label:=New Item","to_class:=JavaMenu","path:=JMenu;JPopupMenu;JPanel;JLayeredPane;JRootPane;PSContentExplorerApplet;PluginEmbeddedFrame;").JavaMenu("label:=MORE","to_class:=JavaMenu","class description:=menu_item")
'		set brMenuList = ParentBrw.JavaMenu("label:=New Item","to_class:=JavaMenu","path:=JMenu;JPopupMenu;JPanel;JLayeredPane;JRootPane;Popup\$HeavyWeightWindow;PluginEmbeddedFrame;")
'		If UCase(More) = "MORE"  Then
'				brMenuList.JavaMenu("label:=MORE","to_class:=JavaMenu","class description:=menu_item").JavaMenu("label:="&MenuName,"to_class:=JavaMenu","index:=0").Select
'			Else
'				brMenuList.JavaMenu("label:="&MenuName,"to_class:=JavaMenu","index:=0").Select
'		End If
'		
'End Function 

''============================================================
'' Function Name: AddContentDetails
''============================================================

Function EnterEphoxContents
		wait 6
		browser("micclass:=browser","index:=1").Page("micclass:=Page","index:=1").JavaApplet("tagname:=EditLiveJava","to_class:=JavaApplet").JavaEdit("tagname:=HTMLPane","to_class:=JavaEdit").Set "Entered from Ephox editor -- " &  arrObjValue(1)
		Wait 6
End Function


''============================================================
'' Function Name: OpenContentMenu
''============================================================
Function OpenContentMenu(FolderName)
	browser("name:=Rhythmyx - Content Explorer").Page("title:=Rhythmyx - Content Explorer").Sync
		wait 1
		If  SiteName = FolderName Then
			LeftPlanel.OpenContextMenu "Content Explorer;Sites;"& SiteName &"" 
		Else
			While LeftPlanel.Exist = False
				Wait 1
			Wend
			Wait 2
			LeftPlanel.OpenContextMenu "Content Explorer;Sites;"& SiteName &";"&FolderName&"" 
'			LeftPlanel.OpenContextMenu "Content Explorer;Sites;"& FolderName&"" 
		End If
End Function



''============================================================
'' Function Name: SelectLeftPlanelMenu
''============================================================
'This is a new function that is capable of selecting the java menu under "New item" in percussion 
Function SelectLeftPanelMenu(MenuName)
	Set ParentBrw = Browser("name:=Rhythmyx - Content Explorer").Page("title:=Rhythmyx - Content Explorer").frame("name:=navcontent").JavaApplet("tagname:=PSContentExplorerApplet","to_class:=JavaApplet")
	set brMenuList = ParentBrw.JavaMenu("label:=New Item","to_class:=JavaMenu","path:=JMenu;JPopupMenu;JPanel;JLayeredPane;JRootPane;Popup\$HeavyWeightWindow;PluginEmbeddedFrame;")
	'Create a description object to select the desired java menu object in a collection object
	Set oDesc = Description.create
	oDesc("to_class").value = "JavaMenu"
	oDesc("class description").value = "menu_item"
	oDesc("toolkit class").value = "javax\.swing\.JMenuItem"
	oDesc("label").value = MenuName

	set oMenuCol = brMenuList.ChildObjects(oDesc)
	MenuCnt = oMenuCol.count
	If MenuCnt = 0 Then
		Msgbox "The menu " & MenuName & " is not found"
	Else
		oMenuCol(0).Select
	End If
		
End Function 

''============================================================
'' Function Name: SelectLeftPlanelMenu
''============================================================
'This is a new function that is capable of selecting the specified java menu under on contents in the right panel in percussion 
Function SelectRightPanelMenu(strMenuName)
	Set ParentBrw = Browser("name:=Rhythmyx - Content Explorer").Page("title:=Rhythmyx - Content Explorer").frame("name:=navcontent").JavaApplet("tagname:=PSContentExplorerApplet","to_class:=JavaApplet")
	set brMenuList = ParentBrw.JavaMenu("label:=New Item","to_class:=JavaMenu","path:=JMenu;JPopupMenu;JPanel;JLayeredPane;JRootPane;Popup\$HeavyWeightWindow;PluginEmbeddedFrame;")
	'Create a description object to select the desired java menu object in a collection object
	Set oDesc = Description.create
	oDesc("to_class").value = "JavaMenu"
	oDesc("class description").value = "menu_item"
	oDesc("toolkit class").value = "javax\.swing\.JMenuItem"
	oDesc("label").value = MenuName

	set oMenuCol = brMenuList.ChildObjects(oDesc)
	MenuCnt = oMenuCol.count
	If MenuCnt = 0 Then
		Msgbox "The menu " & MenuName & " is not found"
	Else
		oMenuCol(0).Select
	End If
		
End Function 

''============================================================
'' Function Name: ChangeWorkflow
''============================================================
Function ChangeWorkflow(ContentName,WorkflowType)

'	 Set CommentBox= ParentBrw.JavaDialog("class description:=window","label:=Enter workflow comment\(Optional\)","path:=JDialog;PluginEmbeddedFrame;","to_class:=JavaDialog","toolkit class:=javax\.swing\.JDialog","index:=0")
	 Set CommentBox = ParentBrw.JavaDialog("class description:=window","title:=Enter Workflow Comment.*")
	 SelectContentFileMenu(ContentName)
	 Wait 1
	 
	 Set oDesc = Description.create
	oDesc("to_class").value = "JavaMenu"
	oDesc("class description").value = "menu_item"
	oDesc("toolkit class").value = "javax\.swing\.JMenuItem"
	oDesc("label").value = WorkflowType
	
	set oMenuCol = ParentBrw.JavaMenu("label:=Workflow","to_class:=JavaMenu").ChildObjects(oDesc)
	MenuCnt = oMenuCol.count
	If MenuCnt = 0 Then
		Msgbox "The menu " & WorkflowType & " is not found"
	Else
		oMenuCol(0).Select
	End If

	 If  CommentBox.Exist(1) = True Then
	 	Commentbox.JavaEdit("path:=JTextArea;JViewport;JScrollPane;JPanel;JLayeredPane;JRootPane;JDialog;PluginEmbeddedFrame;").Set "Approved by the HP UFT Functional testing tool"
		Commentbox.JavaButton("label:=OK").Click
	 End If
	 wait 1
End Function

'Function ChangeWorkflow(ContentName,WorkflowType)
'		
''		 Set CommentBox = ParentBrw.JavaDialog("label:=Enter Workflow Comment.*","to_class:=JavaDialog").JavaButton("label:=OK")
'		 Set CommentBox = ParentBrw.JavaDialog("class description:=window","title:=Enter workflow comment.*")
'		 SelectContentFileMenu(ContentName)
'		 Wait 1
'		 ParentBrw.JavaMenu("label:=Workflow","to_class:=JavaMenu").JavaMenu("label:="&WorkflowType,"to_class:=JavaMenu").Select
'
'		 If  CommentBox.Exist(1) Then
''		 		ParentBrw.JavaDialog("label:=Enter Workflow Comment.*","to_class:=JavaDialog").JavaEdit("tagname:=Please enter comment for the workflow transition\.").Set "Approved by the HP UFT Functional testing tool"
'		 	Commentbox.JavaEdit("path:=JTextArea;JViewport;JScrollPane;JPanel;JLayeredPane;JRootPane;JDialog;PluginEmbeddedFrame;").Set "Approved by the HP UFT Functional testing tool"
'			Commentbox.JavaButton("label:=OK").Click
'		 End If
'		 wait 1
'End Function

Function PreviewContent (strCommunity, strContentType, strContentTitle, strPath, strPreviewTemplate)
'	SetCommunityLanguageFolder strCommunity, strPath
'	WaitForBrowserSync
'	Wait 2
	SelectContentFileMenu(strContentTitle)
	WaitForBrowserSync
'	ParentBrw.JavaMenu("label:=Preview - CancerGov","to_class:=JavaMenu").JavaMenu("label:="&strPreviewTemplate,"to_class:=JavaMenu").Select
	ParentBrw.JavaMenu("path:=JMenu;JPopupMenu;JPanel;JLayeredPane;JRootPane;Popup\$HeavyWeightWindow;PluginEmbeddedFrame;","label:=Preview - CancerGov","to_class:=JavaMenu").JavaMenu("toolkit class:=javax\.swing\.JMenuItem","label:="&strPreviewTemplate,"to_class:=JavaMenu","index:=0").Select
	WaitForBrowserSync
	Msgbox "Verify preview of the content " & strContentTitle & "with the template " & strPreviewTemplate
	With Browser("micclass:=browser","index:=1").Page("micclass:=page","index:=0")
		If .WebElement("html tag:=H2","outertext:=Problem during the assembly of item").Exist(1) = True then
			Msgbox "Assembly error detected for the content " & strContentTitle & "with the template " & strPreviewTemplate
		End if 
		Browser("micclass:=browser","index:=1").Close
	End With
	
End Function

Function CreateTranslation (ContentName, Locale)
	SelectContentFileMenu(ContentName)
	ParentBrw.JavaMenu("label:=Create Translation","to_class:=JavaMenu","toolkit class:=javax\.swing\.JMenuItem","index:=0").Select
	Set ObjBr = Browser("name:=Rhythmyx - New Translation Version Properties").Object
	While ObjBr.readystate <> 4
		Wait 1
	Wend
	
	With Browser("name:=Rhythmyx - New Translation Version Properties").Page("name:=translateitem")
		.Weblist("name:=sys_lang").Select Locale
		.WebButton("name:=Create").Click
	End With
	
	Browser("name:=Rhythmyx - New Translation Version Properties").Dialog("text:=Message from webpage").WinButton("text:=OK").Click
	Browser("name:=Rhythmyx - Translation Results").Page("name:=translateitem").WebButton("name:=Close").Click
	
	'Edit the translated content; Change the title and the pretty url
	
	
End Function

'#####################################################################################
'The following function edits and updates the specified fields of the content; 
'The content should exist in an editable state for this function to work as expected
'#####################################################################################

Function EditContent (strContentTitle, CustomFieldType, CustomValue)
	
	Select Case CustomFieldType
		Case "Title"
			.WebEdit("html tag:=INPUT", "type:=text", "name:=long_title").Set CustomValue
		Case "PrettyURL"
			.WebEdit("html tag:=INPUT", "type:=text", "name:=pretty_url_name").Set CustomValue
	End Select
			
End Function

''============================================================
'' Function Name: SelectContentFileMenu
''============================================================

Function ChangeSiteName(strSiteName)
   Environment.Value("SiteName")= strSiteName
End Function

Function SelectContentFileMenu(ActualContentName)
		Dim ExistValue
		Dim i
		Dim Content_Count
		Dim ContentName
		Dim strTestCaseNumber
		Dim strTestCaseDescription
		
		 Content_Count =RightPlanel.GetROProperty("rows")
		 For i =0 to Content_Count -1
			 ContentName= RightPlanel.GetCellData(i,0)
			 If  Split(ContentName,"[",-1,1)(0) =ActualContentName  Then
				 RightPlanel.ClickCell i,"Content Title","RIGHT","NONE" 
			 End If
		 Next
End Function

Sub ViewContentFileMenu(ActualContentName)
'	   	On Error Resume Next
	
		Dim i
		Dim Content_Count
		Dim ContentName
		
		 Content_Count =RightPlanel.GetROProperty("rows")
		 For i =0 to Content_Count -1
			 ContentName= RightPlanel.GetCellData(i,0)
			 If  Split(ContentName,"[",-1,1)(0) =ActualContentName  Then
				 RightPlanel.ClickCell i,"Content Title","LEFT","NONE" 
			 End If
		 Next
End Sub

''============================================================
'' Function Name: AddContentIntoSlot
''============================================================
Function AddSlotContents(ActualContentName,SlotName,ContentPath,ContentName,strSlotItemType,TemplateName)
		ArrContentName = Split(ContentName, ":")
		ArrContentName = Split(ContentName, "|")
		If Ubound(ArrContentName) >0 Then
			ContentPath1 = ArrContentName(0) & "/" & ArrContentName(1)
			ContentName1 = ArrContentName(2)
		Else
			ContentPath1 = ContentPath
			ContentName1 = ContentName
		End If
		Call SelectContentFileMenu(ActualContentName)
		
		ParentBrw.JavaMenu("label:=Active Assembly Table Editor","to_class:=JavaMenu","index:=0").Select
		With Browser("name:=Related Content Control").Page("title:=Related Content Control")				
			wait 1
			.Link("innertext:=.*"& Replace(Replace(SlotName, "(", "\("), ")", "\)") &".*","index:=0").Click
	    End With
	    wait 1

		
	Browser("micclass:=browser","title:=Content Browser").Page("title:=Content Browser").WebEdit("name:=WebEdit","index:=0").Set Replace(ContentPath1, "|", "/")

	With Browser("micclass:=browser","name:=Content Browser").Page("title:=Content Browser")
		If strSlotItemType = "" Then 'content type is not specified
			.Link("innertext:="&ContentName &".*","index:=0").Click
		Elseif(SlotName =  "nciSectionNavRootNavon")  Or  Instr(SlotName, "PDQ Cancer Information Summary") <> 0 or strSlotItemType = "Custom Link" Then 'content ID is not applicable
			.Link("innertext:="&ContentName1,"index:=0").Click
		Else
			.Link("innertext:=.*"& ContentName1 &"\[.*\]$","index:=0").Click
		End If
		
		If TemplateName <> "" Then
			.WebList("name:=select","select type:=Single Selection").Select TemplateName
		End If
		
		wait 1
		strTemp1 = .WebElement("html id:=ps\.select\.templates\.wgtPreviewPane","html tag:=DIV").GetROProperty("innertext")
		If Instr(strTemp1, "Problem assembling output for item") <> 0 Then
			Msgbox "Assembly Error Detected"
			strTestStep = "Adding the content " & ContentName & " to slot " & SlotName & " with template " & TemplateName
			ResultWriteToFile False, "1", strTestStep, "Content preview with the selected template", strTemp1
		End If
		
		.WebElement("innertext:=Select|Open","html tag:=DIV","index:=0").Click
        wait 1
	End With		
		
	Browser("name:=Content Browser","index:=0").close
	Browser("micclass:=browser","name:=Related Content Control").Close
		
End	Function

Function SelectContentFromContentBrowser (ContentPath,ContentName,strSlotItemType,TemplateName)
	
	Browser("micclass:=browser","name:=Content Browser").Page("title:=Content Browser").WebEdit("name:=WebEdit","index:=0").Set ContentPath

	With Browser("micclass:=browser","name:=Content Browser").Page("title:=Content Browser")
		If(SlotName =  "nciSectionNavRootNavon")  Or  Instr(SlotName, "PDQ Cancer Information Summary") <> 0 or strSlotItemType = "Custom Link" Then 'content ID is not applicable
			.Link("innertext:="&ContentName,"index:=0").Click
		Else
			.Link("innertext:=.*"& ContentName &"\[.*\]$","index:=0").Click
		End If

		.WebList("name:=select","select type:=Single Selection").Select TemplateName
		wait 1
		strTemp1 = .WebElement("html id:=ps\.select\.templates\.wgtPreviewPane","html tag:=DIV").GetROProperty("innertext")
		If Instr(strTemp1, "Problem assembling output for item") <> 0 Then
			Msgbox "Assembly Error Detected"
			strTestStep = "Adding the content " & ContentName & " to slot " & SlotName & " with template " & TemplateName
			ResultWriteToFile False, "1", strTestStep, "Content preview with the selected template", strTemp1
		End If
		
		.WebElement("innertext:=Select","html tag:=DIV","index:=0").Click
        wait 1
	End With		
		
	
End Function

'The following function adds the specified content type to the specified slot with specified template; This is useful when the name of the content to be added is not specified.
Function AddSlotContentsByType(ActualContentName,SlotName,ContentPath,strSlotItemType,TemplateName)

		Call SelectContentFileMenu(ActualContentName)
		
		ParentBrw.JavaMenu("label:=Active Assembly Table Editor","to_class:=JavaMenu").Select
		With Browser("name:=Related Content Control").Page("name:=rcedit")				
			wait 1
			.Link("innertext:=.*"& Replace(Replace(SlotName,"(","\("), ")","\)") &".*","index:=0").Click
	    End With
	    wait 1

	'Set the Path to choose supporting contents from
	With Browser("micclass:=browser","name:=Content Browser").Page("name:=contentBrowerDialog","title:=Content Browser")
		.WebEdit("name:=WebEdit","index:=0").Set ContentPath
'		.Image("name:=Image","alt:=Refresh").Click
	End With
	
	'Get the name of the first available cotent of the specified type:
	Set ObjTable = Browser("micclass:=browser","name:=Content Browser").Page("name:=contentBrowerDialog","title:=Content Browser").WebTable("html tag:=TABLE","class:=ps_content_browse_viewtable","column names:=Name;Desc","html id:=ps\.content\.sitespanel\.FilteringTable")
	Rows = ObjTable.RowCount
	boolContentFound = False
	For idx2 = 2 To Rows
		ContDescription = ObjTable.GetCelldata(idx2,2)
		If ContDescription = strSlotItemType Then
			ContentName = Replace(Replace(Trim(ObjTable.GetCelldata(idx2,1)), "[", "\["), "]","\]")
			boolContentFound = True
			Exit For
		End If
	Next 
	If boolContentFound = False Then
		Msgbox "The Specified content type:" & strSlotItemType & "doesn't exist in the path: " & ContentPath 
		wait 1
	End If

	With Browser("micclass:=browser","name:=Content Browser").Page("name:=contentBrowerDialog","title:=Content Browser")
		.Link("innertext:=" &Replace(Replace(ContentName,"(","\("), ")","\)"),"index:=0").Click
		.WebList("name:=select","select type:=Single Selection").Select TemplateName
		wait 1
		strTemp1 = .WebElement("html id:=ps\.select\.templates\.wgtPreviewPane","html tag:=DIV").GetROProperty("innertext")
		strTestStep = "Adding the content " & ContentName & " to slot " & SlotName & " with template " & TemplateName
		If Instr(strTemp1, "Problem assembling output for item") <> 0 Then
			Msgbox "Assembly Error Detected"
			ResultWriteToFile False, "1", strTestStep, "Content preview with the selected template", strTemp1
		End If
		
		.WebElement("innertext:=Select","html tag:=DIV","index:=0").Click
        wait 1
	End With		
	ResultWriteToFile True, Environment.value("strTestCaseNumber"), strTestStep, "Add the content with selected template", "Step completed successfully"	
	Browser("name:=Content Browser","index:=0").close
	Browser("micclass:=browser","name:=Related Content Control").Close
		
End	Function




''============================================================
'' Sub Name: SelectCommunity
''============================================================		
Sub SelectCommunity ()
					   
browser("micclass:=browser","index:=0").Page("micclass:=Page","index:=0").Link("html tag:=A","index:=6").click

Select Case UCase(Environment.Value("SiteName"))
	Case "CGOV"
		browser("micclass:=browser","index:=0").Page("micclass:=Page","index:=0").WebList("html tag:=SELECT","name:=community").Select "CancerGov"
	Case "TCGA"	
		browser("micclass:=browser","index:=0").Page("micclass:=Page","index:=0").WebList("html tag:=SELECT","name:=community").Select "TCGA"
	Case "IMAGING"	
		browser("micclass:=browser","index:=0").Page("micclass:=Page","index:=0").WebList("html tag:=SELECT","name:=community").Select "Imaging"
	Case "PROTEOMICS"	
		browser("micclass:=browser","index:=0").Page("micclass:=Page","index:=0").WebList("html tag:=SELECT","name:=community").Select "Proteomics"
End Select

browser("micclass:=browser","index:=0").Page("micclass:=Page","index:=0").WebList("html tag:=SELECT","name:=sys_lang").Select "US English"
browser("micclass:=browser","index:=0").Page("micclass:=Page","index:=0").Image("file name:=go.gif","html tag:=INPUT").Click
wait 4
End Sub

''============================================================
'' Sub Name: ClickOnEditContentLink
''============================================================
Sub ClickOnEditContentLink(ActualContentName)		
		
		Call SelectContentFileMenu(ActualContentName)
		wait 1
		ParentBrw.JavaMenu("label:=Edit","to_class:=JavaMenu").Select
End Sub
		
'Function ActiveAssemblyTable(SlotName,ContentName)
			
'			ParentBrw.JavaMenu("label:=Active Assembly Table Editor","to_class:=JavaMenu").Select
'			wait 2
'			Call AddContentIntoSlot(SlotName,ContentName)
'End Function		

''============================================================
'' Function Name: SiteName
''============================================================
Function SiteName()
			Select Case Lcase(Environment.Value("SiteName"))
				Case "stage"
					strTmp = "-stage.cancer.gov"
				Case "qa"
					strTmp = "-qa.cancer.gov"
				Case "blue"
					strTmp = "-blue-dev.cancer.gov"
				Case "red"
					strTmp = "-red.dev.cancer.gov"
				Case "prod", "production"
					strTmp = ".cancer.gov"
				Case "dt"
					strTmp = "-dt-qa.cancer.gov"
				Case "pink"
					strTmp = "-pink-dev.cancer.gov"
			End Select
			
			 Select Case UCase(Environment.Value("SiteName"))
					Case "CANCERGOV", "CGOV", "CANCERGOV_CONFIGURATION"
						 SiteName ="CancerGov"
						 Check = True
						 Environment.Value("FirstURLPart") = "www" & strTmp
						 
					Case "TCGA"
						 SiteName ="TCGA"
						 Check = True
						 Environment.Value("FirstURLPart") = "cancergenome" & strTmp '"tcgaint.cancer.gov"
						 
					Case "MOBILE","MOBILECANCERGOV"
						 SiteName ="MobileCancerGov"	
						 Check = True	
						 Environment.Value("FirstPart") = "m" & strTmp
						 						 
					Case "PROTEOMICS", "PROTEOMICS_CONFIGURATION"
						 SiteName ="Proteomics"
						 Check = True
						 Environment.Value("FirstPart") = "proteomics" & strTmp
						 
					Case "IMAGING", "IMAGING_CONFIGURATION"
						Environment.Value("FirstURLPart") = "imaging" & strTmp
						 SiteName ="Imaging"
						 Check = True

					Case "CCOP", "IMAGING_CONFIGURATION"
						SiteName ="CCOP"
						Environment.Value("FirstURLPart") = "ccop" & strTmp

					Case "DCEG", "IMAGING_CONFIGURATION"
						SiteName ="DCEG"
						Environment.Value("FirstURLPart") = "dceg" & strTmp

					Case Else
						Msgbox "Please add a valid entery for " & UCase(Environment.Value("SiteName")) & "in the case statement"
			End Select	

			
End Function		

''============================================================
'' Sub Name: WorkflowValidation
''============================================================

'Function WorkflowValidation(ActualContentName)
Function CheckWorkflowStatus(strContentTitle, strContentType, strExpectedStatus)	
		RefreshContentExplorer() 
		WaitForBrowserSync
		 Content_Count =RightPlanel.GetROProperty("rows")
		 For i =0 to Content_Count -1
			ContentName= RightPlanel.GetCellData(i,0)
			RuntimeContentType = RightPlanel.GetCellData(i,2)
			RuntimeContentName = Split(ContentName, "[", -1, 1)
			If  RuntimeContentName(0) = strContentTitle And RuntimeContentType = strContentType Then
				 Wflstatus= RightPlanel.GetCellData(i,3)
				 boolResult =  (Wflstatus = strExpectedStatus)
				 Exit For
			 End if
		 Next
		 strTestStep = "Comparing the workflow status of the content " & strContentTitle 
		 ResultWriteToFile boolResult, "1", strTestStep, strExpectedStatus, Wflstatus
		 If boolResult =False Then Msgbox "Content " & strContentTitle & " not in the right workflow status" End If
		 CheckWorkflowStatus = boolResult
End Function 

Function ReportResults(ActualContentName)
'   On error Resume Next
   wait 1
		Dim Returndata
		Dim PageName
		object = ActualContentName
		PageName = datatable.value(4)
'		 If ActualContentName = True  Or ActualContentName = PageName then
		If  ActualContentName = PageName then	'Validating Page contents using reportresults function
				Call GetContentID(ActualContentName)
				Returndata = FrontEndValidation(PageName)
				if ReturnData <> "" Then
					strObject = "Fail"
					strActRslt = ReturnData
				End if 
		 ElseIF ActualContentName = False  Then	'Validating the suporting contents using outputresults
			  strObject = "Fail"
			  strActRslt = ActualContentName &" Content File not found or not in right workflow staus"
		End If	
		If 	strActRslt = "" then
			strActRslt = "Verified and All the Contents are displaying successfully."
			strObject = "Pass"
		End If 
		strTestCaseNumber = Environment.value("tcStepNum")  ''''''Environment.Value("CurTestCase#")
		strTestCaseDescription =  Cstr(Trim(DataTable.Value(3)))
    		Call WriteToFile(strObject, strTestCaseNumber, strTestCaseDescription,object,Parentbrw,strActRslt)
			
strResult = strObject
End Function


Function ReadTextFromNotepad (strCompletePath) ' eg:"C:\Temp\FormText.txt"
   
		Const conForReading = 1
		
		'Declare variables
		Dim objFSO, objReadFile, contents
		
		'Set Objects
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objReadFile = objFSO.OpenTextFile(strCompletePath, 1, False)
		
		contents = objReadFile.ReadAll
		
		'Close file
		objReadFile.close
				
		'Cleanup objects
		Set objFSO = Nothing
		Set objReadFile = Nothing
		ReadTextFromNotepad = contents
End Function
