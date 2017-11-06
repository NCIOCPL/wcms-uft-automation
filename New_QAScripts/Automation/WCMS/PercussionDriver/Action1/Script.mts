

'
'With Browser("Rhythmyx - Content Explorer").Page("Rhythmyx - Content Explorer_3").Frame("navcontent").JavaApplet("PSContentExplorerApplet").JavaDialog("Enter workflow comment(Optiona")
'	.JavaEdit("Please enter comment for").SetCaretPos 0,0
'	.JavaEdit("Please enter comment for").Set "hello"
''End With
'
'Set ParentBrw = Browser("name:=Rhythmyx - Content Explorer").Page("title:=Rhythmyx - Content Explorer").frame("name:=navcontent").JavaApplet("tagname:=PSContentExplorerApplet","to_class:=JavaApplet") 
''
''	 Set CommentBox = ParentBrw.JavaDialog("class description:=window","title:=Enter Workflow Comment.*")
'Set CommentBox= ParentBrw.JavaDialog("class description:=window","label:=Enter workflow comment\(Optional\)","path:=JDialog;PluginEmbeddedFrame;","to_class:=JavaDialog","toolkit class:=javax\.swing\.JDialog","index:=0")
'
''	 If  CommentBox.Exist(1) Then
'	 	Commentbox.JavaEdit("path:=JTextArea;JViewport;JScrollPane;JPanel;JLayeredPane;JRootPane;JDialog;PluginEmbeddedFrame;").Set "Approved by the HP UFT Functional testing tool"
'		Commentbox.JavaButton("label:=OK").Click
''	 End If
'	 wait 1
'strhtmlID = "ce_bodyfieldbodyfield_ifr"
'	Set brContInfoPg = browser("micclass:=browser","name:=Rhythmyx .*Edit Content").Page("micclass:=Page","title:=Rhythmyx .*Edit Content")
'	Set oDesc = Description.Create()
'	oDesc("micclass").Value = "WebElement"
'	oDesc("html tag").Value = "P"
'	Set oEditArea = brContInfoPg.Frame("html tag:=IFRAME", "html id:="&strhtmlID & ".*|ce_" & strhtmlID & ".*").ChildObjects(oDesc)	
'	oEditArea(0).drag
'	oEditArea(0).drop
'	
'	
'	'Create device replay object
'	 Set oDR = CreateObject("Mercury.DeviceReplay")
'	'Send some text
'	 oDR.SendString "OHHHH LA LA LA OHHHH LA LA LA OHHHH LA LA LA OHHHH LA LA LA OHHHH LA LA LA"
'	 Wait 1	 
'	 

	
'Variable Declarations
Dim TestCasePath
Dim gblUrl
Dim strTempVar1
Dim TestcaseFile
Dim intIDx
UserName =  Environment.Value("UserName")
LocalHostName = Environment.Value("LocalHostName")
Environment.Value("Environment") = "dt"

Dim ArgumentList


TestcaseFile = Inputbox ("Please enter the testcase file name")
If TestcaseFile = "" Then
	TestcaseFile = "Driver_Cgov_Contents.xls"'"Driver_TCGA_Contents.xls" '"Driver_DCEG_Contents.xls" '"Driver_Cgov_CgovConfig_Contents.xls"  Driver_TCGA_Contents, '"Driver_Cgov_CgovConfig_Contents.xls"
End If

Environment.Value("TestcaseFile") = Replace(TestcaseFile, ".xls", "")
'#################################################################################################
'ParentPath ="L:\OCPL\ODDC\CPSB\QAStuff\Veena\Automation\"
'New ParentPath
ParentPath ="L:\OCPL\ODDC\CPSB\QAStuff\New_QAScripts\Automation\"
       
Environment.Value("ResultsPath") = ParentPath & "Results\WCMS\Percussion\"
Environment.Value("ImageResPath")= ParentPath & "Results\WCMS\Percussion\Images\" 
Environment.value("strTestCaseNumber")=1 'Default value
'The followingline is commented for data files on L:drive and is provided a link on the local C drive because, it was taking too long for the windows explorer to locate files deep down the L drive heirarchy and scripts failed often.

'Environment.Value("TestDataPath") = ParentPath & "Projects\WCMS\Datasheets\Data_files\"
Environment.Value("TestDataPath") = "C:\Automation\Data_files\"
Environment.Value("TestContentPath") = "\Automation_RealWorldScenarios\"
'TestCasePath = ParentPath & "Projects\WCMS\Datasheets\"
'New TestCasePath
TestCasePath = ParentPath & "DataSheetsForPercussion\"
'#################################################################################################
 

Datatable.Import  TestCasePath & TestcaseFile
intRowCount = Datatable.GetRowCount
For intIDx=1 to intRowCount 
    DataTable.SetCurrentRow (intIDx) 
	If  DataTable("Status", dtGlobalSheet) <> "Done" Then
		 Environment.value("strTestCaseNumber")=intIDx
		Environment.value("tcStepNum")=intIDx
		strScriptName = Datatable.Value("Testcase", dtGlobalSheet)
		strArguments = Datatable.Value("Arguments", dtGlobalSheet)
		ArrTestcase = Split(strScriptName, ",")
		ArrArguments = Split(strArguments, ",")

		If strScriptName <> "" Then
			If ubound(ArrArguments)>1 then strContentTitle = Trim(ArrArguments(0)) End if
'			strContentTitle = Trim(ArrArguments(0)) 
			strSiteName = Trim(ArrTestcase(1))
			If Ubound(ArrArguments)>1 Then
				strContentPath = Trim(ArrArguments(1))
			End If
			Environment.Value("strTestName") = strSiteName
			If Ubound(ArrTestcase)>2 Then
				strOtherOptions = Trim(ArrTestcase(3))
			Else
				strOtherOptions = ""
			End If
			 
			Select Case Trim(ArrTestcase(0))
				Case "CreateContent", "PreviewContent" 					
					strPath = Trim(ArrArguments(1))
					strStatus = Trim(ArrArguments(2))
					strCommunity = Trim(ArrTestcase(1))
					strContentType = Trim(ArrTestcase(2))
					Environment.Value("strTestName") = "CreateContent - " & ArrTestcase(2)
					Environment.Value("SiteName") = Replace(strCommunity, "_Configuration", "") 
					If Trim(ArrTestcase(0)) = "CreateContent" Then
						If Ubound(ArrTestcase) >2 Then
							If Trim(ArrTestcase(3)) = "FillOnlyRequiredFields" Then
								FillOnlyRequiredFields = True
							End If
						Else
							FillOnlyRequiredFields = False
						End If
						
						CreateContent strCommunity, strContentType, strContentTitle, strOtherOptions, strPath, FillOnlyRequiredFields
					ElseIf Trim(ArrTestcase(0)) = "PreviewContent" Then
						strPreviewTemplate = Trim(ArrArguments(2))
						PreviewContent strCommunity, strContentType, strContentTitle, strPath, strPreviewTemplate
					End If
					
				If lcase(strStatus) = "public" Then
					If ubound(ArrArguments) = 3 then strFlowType = Trim(ArrArguments(3)) Else strFlowType = "" End If'Regular or CDR
						PushWorkFlowFromDraftToPublic strContentTitle, strFlowType
					End If
				
				Case "VerifySyndicatedPage"
					Msgbox "Go to publishing runtime in Percussion and run Syndication Live job"
					Msgbox "Verify the Syndicated page at Live/publishedcontent/syndication/ContentID.htm And (Cancergov version) at Live/Automation_Syndication/prettyurl"
					
					
				Case "AddSlotContents"
					Environment.Value("strTestName") = "AddContentsToSlots -  " & ArrTestcase(1)		
					strSlotName = Trim(ArrArguments(2))
					strSlotItemName = Trim(ArrArguments(3))
					If ubound(ArrArguments) >4 Then strSlotItemType = Trim(ArrArguments(5)) Else strSlotItemType = "" End IF
					If ubound(ArrArguments) >3 Then TemplateName = Trim(ArrArguments(4)) Else TemplateName = "" End If
					strEditorOptions = Ucase(Trim(ArrTestcase(1)))
					AddSlotContents strContentTitle, strSlotName, "/" & strSiteName & "/" & strContentPath, strSlotItemName, strSlotItemType, TemplateName
				
				Case "PushWorkFlowFromDraftToPublic"
					strFlowType = Ucase(Trim(ArrArguments(2))) 'Regular or CDR
					strPath = Trim(ArrArguments(1))
					strCommunity = Trim(ArrTestcase(1))
					Environment.Value("SiteName") = Replace(strCommunity, "_Configuration", "") 
					SelectCommunityLanguage strCommunity, "US English"
					SelectContentFolder strPath
					PushWorkFlowFromDraftToPublic strContentTitle, strFlowType
					
				Case "CheckWorkflowErros", "ChangeWorkflow"
					strCommunity = Trim(ArrTestcase(1))
					strPath = Trim(ArrArguments(1))
					strContentType = Trim(ArrTestcase(2))
					Environment.Value("SiteName") = Replace(strCommunity, "_Configuration", "") 
					strNewWorkFlow = Trim(ArrArguments(2))
					If Ubound(ArrArguments) > 2 Then ErrorOptions = Trim(ArrArguments(3)) Else ErrorOptions = "" End If
					SelectCommunityLanguage strCommunity, "US English"
					SelectContentFolder strPath
					CheckWorkflowErros strContentTitle, strNewWorkFlow, ErrorOptions
					
				Case "PageSlotValidation"
					
					CompareContentFrontendWithBackend strContentTitle
				
				Case "CheckTemplateErrors"
					strContentType = Trim(ArrTestcase(2))
					arrSlotNames = Split(ArrArguments(2), ":")
					For i = 0 To Ubound(arrSlotNames)	
						strSlotName = arrSlotNames(i)
						CheckTemplateErrors strSiteName, strContentPath, strContentTitle, strContentType, strSlotName
					Next
					
				Case "CheckTemplateErrorsAllSlots"
					strContentPath = Trim(ArrArguments(1))
					strContentType = Trim(ArrTestcase(2))
					strContentTitle = Trim(ArrArguments(0))
					CheckTemplateErrorsAllSlots strSiteName, strContentPath, strContentTitle, strContentType
				
				'CreateFolderContent, CancerGov, AutomationScenarios
				Case "CreateFolderContent"
					strCommunity = Trim(ArrTestcase(1))
					Environment.Value("SiteName") = Replace(strCommunity, "_Configuration", "") 
					strPath = Trim(ArrArguments(0)) 
					arrTemp = Split(strPath, ";")
					
					If  Ubound(arrTemp) > 0 Then 
						cnt = Ubound(arrTemp)
						NewFolderName = Trim(arrTemp(cnt)) 
'						OldFolderName = Environment.Value("SiteName")
						For Iterator = 0 To cnt-1
							OldFolderName = OldFolderName & ";" & Trim(arrTemp(Iterator))
						Next
						OldFolderName = Right(OldFolderName, Len(OldFolderName)-1) 'Removes the leading ";"
					Else 
						NewFolderName = Trim(arrTemp(0))
						OldFolderName = Environment.Value("SiteName")
					End IF
					SetCommunityLanguageFolder strCommunity, OldFolderName
					Status = Trim(ArrArguments(1))
					CreateFolderContent OldFolderName, NewFolderName, Status
				
				Case "BuildPageScenario"
					strCommunity = Trim(ArrTestcase(1))
					strPath = Trim(ArrTestcase(2))
					strContentType = Trim(ArrTestcase(3))
					strContentTitle = Trim(ArrTestcase(4))
					
					If Ubound(ArrTestcase) >=5 Then strFlowType = Trim(ArrTestcase(5)) Else strFlowType = "Regular"
					Environment.Value("SiteName") = Replace(strCommunity, "_Configuration", "") 
					'First create the content in the draft status
					'The following line is temporarily commented; Uncomment once the "AddSlotContents" function starts working;
					CreateContent strCommunity, strContentType, strContentTitle, "", strPath, True 'FillOnlyRequiredFields = True
					
					'Parse the slot-snippetTemplate-contenttype delimited ("+") array and add one iten at a time to the specified slot.
					strArguments = Right(strArguments, Len(strArguments)-3)
					ArrSlotTemplateContentInfo = Split(strArguments,"|")
					
					If Environment.Value("SiteName") = "Imaging" Or Environment.Value("SiteName") = "Proteomics" Or Environment.Value("SiteName") = "CCOP" Then
						strTempPath = "DCEG"
					Else
						strTempPath = Environment.Value("SiteName")	
					End If
					'Add the slot contents one at a time
					For idx=0 To Ubound(ArrSlotTemplateContentInfo)
						ArrSingleSlotInfo = Split(ArrSlotTemplateContentInfo(idx), "+")
						Environment.Value("strTestName") = "AddContentsToSlots -  " & strContentTitle
						strSlotName = Trim(ArrSingleSlotInfo(0))
						TemplateName = Trim(ArrSingleSlotInfo(1))
						strSlotItemType = Trim(ArrSingleSlotInfo(2))										
						AddSlotContentsByType strContentTitle,strSlotName, "/" & strTempPath & "/Automation",strSlotItemType, TemplateName
					Next
					
					'Push the content workflow status from Draft to Public
					PushWorkFlowFromDraftToPublic strContentTitle, strFlowType	
			
				Case "SelectCommunityLanguage"
					strCommunity = Trim(ArrTestcase(1))
					strLanguage = Trim(ArrTestcase(2))
					SelectCommunityLanguage strCommunity, strLanguage	
				Case "CheckWorkflowStatus"
					strCommunity = Trim(ArrTestcase(1))
					Environment.Value("strTestName") = "CreateContent - " & ArrTestcase(2)
					Environment.Value("SiteName") = Replace(strCommunity, "_Configuration", "")
					strContentType = Trim(ArrTestcase(2))
					strFlowType = Trim(ArrArguments(2))
					strPath = Trim(ArrArguments(1))
					SelectContentFolder strPath
					CheckWorkflowStatus strContentTitle, strContentType, strFlowType
			End Select 		
		End If

		DataTable.Value("Status", dtGlobalSheet) = "Done"
		DataTable.Export (TestCasePath & TestcaseFile)
	End If	
Next
