
 @@ hightlight id_;_2071267240_;_script infofile_;_ZIP::ssf7.xml_;_
''Dim arrSlots()
'Set DescSlotName = Description.Create()
'DescSlotName("micclass").Value="WebElement"
'DescSlotName("html tag").Value= "DIV" 
'DescSlotName("html id").Value= "cgv.*"
'   
'Set SlotItemList = Browser("micclass:=browser").Page("micclass:=Page").ChildObjects(DescSlotName)
'tmp1=  SlotItemList.count
'For i1 = 0 To tmp1-1 
'	arrSlots(i1) = SlotItemList(i1).GetRoProperty("html id")
'Next
'
'msgbox ubound(arrSlots)
'####################################################################################
'Updated on 08/26/2013 by Veena Gubbi
'The script is now configured to pull slot data dynamically from the browser
'User need not have to plug in the slot info for screen scraping now
'This takes care of any new slots added/old slots dropped
'======================================================================================
' Variable Declaration 
'======================================================================================
'
'L:\OCE\CTB-AppDev\QAStuff\Veena\Automation\Projects\WCMS\DataSheetsScreenscrape
 Dim Results, ResultsPath,FirstRow
 Dim RowCount, RunFrom, RunTo,GoToPage
 Dim Mytim,CurrentMonth,CurrentDate,CurrentDay,Hours
 Hours=Hour(Now)
 Minutes=Minute(Now)
 CurrentMonth= Month(date)
  CurrentDay= Day(Date)
  Seconds=Second(Now)
  ParentPath = "L:\OCPL\ODDC\CPSB\QAStuff\New_QAScripts\Screen-Scrape\"
CurrentDate=CurrentMonth&"_"&CurrentDay 
  FirstRow=1
OutputPath= "L:\OCE\CTB-AppDev\QAStuff\Veena\Automation\Results\WCMS\ScreenScrape\"
 Runtime="_"&CurrentDate&"_"&Hours&"_"&Minutes&"_"&Seconds
CurrentRow=FirstRow
SheetName="DataSheet"


Dim strSelectFunction:strSelectFunction = "empty"
Dim strSiteName:strSiteName = "empty"
Dim boolWaiting
Dim strURLFile
Dim URLFirstPart
Dim CreationTime, Mask
'======================================================================================
																		' ' Main scripts
'======================================================================================
	boolWaiting = True
	

'While boolWaiting = True
	strSelectFunction = InputBox ("Enter 1 to call Before Screen Scrape OR 2 to call After Screen Scrape", "Choose a Function")
	
	strSiteName = UCase (InputBox ("CGOV; CGOV1;CGOV2; CGOV3; TCGA; DCEG;  PLEASE ENTER ONE OF THESE SITES"))
	strEnvironment = UCase (InputBox ("Please enter the environment: qa, dt, red, blue, pink, prod, stage, dev6, dev15 etc.,"))
	If strEnvironment = "" Then
		strEnvironment = "dt"
	End If
	strEnv = strEnvironment
	strEnvironment = Replace(SetSiteURL(strEnvironment), "www", "")
	If  (strSelectFunction = "1" or strSelectFunction = "2") And (strSiteName="DCEG" or Instr(strSiteName, "CGOV") <> 0 or  strSiteName="PROTEOMICS" or strSiteName="CCOP" or strSiteName="IMG" or strSiteName="MOBILE" or strSiteName="TCGA") Then
		boolWaiting = False
	Else
		Msgbox "You have entered  invalid inputs: Please Enter Valid Choices"
	End If	

'============ Main Program=======================

If   strSelectFunction = "1" Then
'Please set the URLFirstPart for all the following sites to appropriate environment for pre-migration screenscrapes
				Select Case UCase(strSiteName) 'These tables are for regular deployment testing; Comment the block above and uncomment this block
					Case "CGOV", "CGOV1", "CGOV2"
						If strSiteName = "CGOV" Then
							strURLFile = "Cgov_URL.xls"
						ElseIf strSiteName = "CGOV1" Then
							strURLFile = "Cgov1_URL.xls"
						ElseIf strSiteName = "CGOV2" Then
							strURLFile = "Cgov2_URL.xls"
						End If
						TableName="Cgov_Ref"
						URLFirstPart = "http://www" 
					Case "TCGA"		
						strURLFile = "TCGA_URL.xls"
						TableName="TCGA_Ref"
						If lcase(strEnv) = "dev13" Then
							URLFirstPart = "http://tcga" 
						Else
							URLFirstPart = "http://cancergenome" 
						End If
'					Case "PROTEOMICS"
'						strURLFile = "Proteomics_URL.xls"
'						TableName="Proteomics_Ref"
'						URLFirstPart = "proteomics"
'					Case "CCOP"
'						strURLFile = "Ccop_URL.xls"
'						TableName="CCOP_Ref"
'						URLFirstPart = "ccop"
'					Case "IMG"
'						strURLFile = "Imaging_URL.xls"
'						TableName="Imaging_Ref"
'						URLFirstPart = "imaging" 
'					Case "MOBILE"
'						strURLFile = "Mobile_URL.xls"
'						TableName="CgovMobile_Ref"
'						URLFirstPart = "m" 
					Case "DCEG"
						strURLFile = "DCEG_URL.xls"
						TableName="DCEG_Ref"
						URLFirstPart = "dceg" 
				End Select
				URLFirstPart = URLFirstPart & strEnvironment
	
Else 
'Please set the URLFirstPart for all the following sites to appropriate environment for post-migration screenscrapes
          		Select Case UCase(strSiteName) 'These tables are for regular deployment testing; Comment the block above and uncomment this block
					Case "CGOV", "CGOV1", "CGOV2"
						If strSiteName = "CGOV" Then
							strURLFile = "Cgov_URL.xls"
						ElseIf strSiteName = "CGOV1" Then
							strURLFile = "Cgov1_URL.xls"
						ElseIf strSiteName = "CGOV2" Then
							strURLFile = "Cgov2_URL.xls"
						End If
						TableName="Cgov_Runtime"
						URLFirstPart = "https://www" 
						
					Case "TCGA"		
						strURLFile = "TCGA_URL.xls"
						TableName="TCGA_Runtime"
						If lcase(strEnv) = "dev13" Then
							URLFirstPart = "http://tcga" 
						Else
							URLFirstPart = "http://cancergenome" 
						End If
						
'					Case "PROTEOMICS"
'						strURLFile = "Proteomics_URL.xls"
'						TableName="Proteomics_Runtime"
'						URLFirstPart = "proteomics"
'					Case "CCOP"
'						strURLFile = "Ccop_URL.xls"
'						TableName="CCOP_Runtime"
'						URLFirstPart = "ccop"
'					Case "IMG"
'						strURLFile = "Imaging_URL.xls"
'						TableName="Imaging_Runtime"
'						URLFirstPart = "imaging" 
'					Case "MOBILE"
'						strURLFile = "Mobile_URL.xls"
'						TableName="CgovMobile_Runtime"
'						URLFirstPart = "m" 
					Case "DCEG"
						strURLFile = "DCEG_URL.xls"
						TableName="DCEG_Runtime"
						URLFirstPart = "dceg" 
				End Select
				URLFirstPart = URLFirstPart & strEnvironment

				
  End if 
  TableName2 = "Href_" & TableName
  
'###################################################################################################################
If strSiteName = "TCGA" And strEnvironment = "prod" Then
	URLFirstPart = Replace(URLFirstPart, ".cancer.gov", ".nih.gov")
End If

DataPath = ParentPath & "Datasheets_ScreenScrape\" & strURLFile
	DataTable.Import (DataPath) ' importing the Excel into datatable 
	
	RowCount=Datatable.GetSheet("TestCase").GetRowCount
  	set conn=Createobject("ADODB.Connection")
  	
	'"Enter your connection string here"
	SITConnectionStr = "***REMOVED***"
	conn.open SITConnectionStr ' connecting to db	
'	RunFromRow = Cint(Datatable.Value("RunFromRow","TestCase") 	)
With Browser("micclass:=browser")
	For RunFrom=1 to RowCount
		Datatable.GetSheet("TestCase").SetCurrentRow(RunFrom)
		strTemp1 = DataTable("Status")
		PrettyURL=Datatable.Value("PrettyURL","TestCase")
		If strTemp1 <> "Done" And PrettyURL <> "" Then
			print "Current URLNumber = " & RunFrom & "/ Total URL Count = " &  RowCount	
			GoToPage = URLFirstPart & PrettyURL
			Browser("micclass:=browser", "index:=0").Navigate GoToPage

			If .WinObject("nativeclass:=window","visible:=True","acc_name:=Notification bar").Exist(0) = True  Then
				.WinObject("nativeclass:=window","visible:=True","acc_name:=Notification bar").WinButton("object class:=push button","acc_name:=Cancel").Click
				Else

					WaitForBrowserSync
					'get all contentID, SlotID,Slotname from each page and store into DB
					Call InsertContentID_into_DB(strSiteName, PrettyURL,TableName,conn, strSelectFunction) ' Insert contenID into DB
					Call InsertLinksAndHrefs_into_DB(strSiteName, PrettyURL,TableName2,conn, strSelectFunction, RunFrom)
			End If

			DataTable.Value("Status", dtGlobalSheet) = "Done"
			DataTable.Export (DataPath)
		End If
	Next
End With
conn.Close
	Set conn=Nothing 
'###################################################################################################################
	'=======Call the function to insert data in the Runtime tables====
	''''Call    Content_Validation(TableName)

 ' ==========Stores the Results into Excel 		=========
''''' DataTable.Export(OutputPath&"results"&Runtime&".xls")	
	

''''=====================================END==============================================
