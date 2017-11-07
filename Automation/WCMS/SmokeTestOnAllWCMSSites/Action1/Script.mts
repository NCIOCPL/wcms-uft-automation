

strEnv = Inputbox ("Please enter one of these Environments: dev13, dev14, dev15, blue red pink qa dt stage prod")
boolRunAppModuleTests = True
If strEnv = "" Then strEnv = "dt" End If

strMode = Inputbox ("Please enter 1 for LIVE, 2 for PREVIEW and 3 for BOTH")
If strMode = "" Then
	strMode = 1
End If

Environment.Value("ResultsPath") = "L:\OCPL\ODDC\CPSB\QAStuff\New_QAScripts\Automation\Results\WCMS\SmokeTest"
                                   
Environment.Value("TestcaseFile") = "WCMS_SmokeTestingReport"
Environment.Value("strTestName") = "WCMS Smoke Test"
Public intTestStep	
intTestStep = 1
Set ParentBr = Browser("micclass:=browser").Page("micclass:=page")

Select Case Lcase(strEnv)
	Case "stage"
		strTmp1 = "-stage.cancer.gov"
	Case "stage1"
		strTmp1 = "-stage1.cancer.gov"
	Case "qa"
		strTmp1 = "-qa.cancer.gov"
	Case "blue"
		strTmp1 = "-blue-dev.cancer.gov"
	Case "red"
		strTmp1 = "-red-dev.cancer.gov"
	Case "prod", "production"
		strTmp1 = ".cancer.gov"
	Case "dt"
		strTmp1 = "-dt-qa.cancer.gov"
	Case "pink"
		strTmp1 = "-pink-dev.cancer.gov"
	Case "training"
		strTmp1 = "-training.cancer.gov"
	Case "dev13"
		strTmp1 = ".ocdev13.ha2.cancer.gov"
	Case "dev14"
		strTmp1 = ".ocdev14.ha2.cancer.gov"
	
	Case "dev15"
	    strTmp1 = "-ocdev15.ha2.cancer.gov"
	
		
End Select
If strEnv="dev15" Then
	ArrSites = Array("https://www","https://cancergenome","https://dceg","https://proteomics","preview","cancergenomepreview","dcegpreview","proteomicspreview")
else
	ArrSites = Array("www","cancergenome","dceg","preview","cancergenomepreview","dcegpreview")
End If			

If strMode = 1 Then
	startIdx = 0
	EndIdx = 2
ElseIf strMode = 2 Then
	startIdx = 3
	EndIdx = 5
ElseIf strMode = 3 Then
	startIdx = 0
	EndIdx = 5
End If

For idx =startIdx To EndIdx 'Ubound(ArrSites)
	If Ucase(ArrSites(idx)) = "M" Then
		Environment.Value("SiteName") = "MOBILE"
	ElseIf Ucase(ArrSites(idx)) = "WWW" Then
		Environment.Value("SiteName") = "CGOV"
	Else
		Environment.Value("SiteName") = Ucase(ArrSites(idx))
	End If
	
	
	If instr(ArrSites(idx), "cancergenome")<>0 And Lcase(strEnv) = "prod" Then
		strTmp = ".nih.gov"
	Else
		strTmp = strTmp1
	End If
	strURL = ArrSites(idx) & strTmp
	
	Browser("micclass:=Browser","index:=0").Navigate strURL 
	WaitForBrowserSync
	
	'####################### Page check and result logs #############################
	strActual = CheckForErrorOnPage
	If Len(strActual) > 0 Then 
		boolResult = False
'		Msgbox CheckForErrorOnPage  
	Else
		boolResult = True
		strActual = "Successful"
	End If
	strTestStep = "Checking the home page " & strURL
	strExpected = "Page loads successfully"

	ResultWriteToFile boolResult, intTestStep, strTestStep, strExpected, strActual
	intTestStep = intTestStep + 1
	'####################### Page check and result logs #############################
					
	With Browser("micclass:=browser").Page("micclass:=page")
		'Checking the header and the major tabs of all WCMS sites; Also checking for any broken pages or error messages on the page
		Set oLinksDefn = Description.Create()
		oLinksDefn("micclass").Value = "Link"
		oLinksDefn("html tag").Value = "A"
		
		Select Case ArrSites(idx)
			Case "www", "www-new", "preview"	
				
				If Len(CheckForErrorOnPage) = 0  Then
					'################# Check the sitewide search functionality ENGLISH ##############
					.WebEdit("class:=searchString ui-autocomplete-input", "name:=swKeyword", "html tag:=INPUT").Set "Cancer"
					.Webbutton("class:=searchSubmit","name:=Search ","type:=submit","html tag:=BUTTON").Click
					WaitForBrowserSync
					boolExists = .WebElement("class:=searchResultsHeader","html tag:=DIV","innertext:=Results 1-10 of.*").Exist(0)
					If boolExists = True Then
						strActual = .WebElement("class:=searchResultsHeader","html tag:=DIV","innertext:=Results 1-10 of.*").GetROProperty("innertext")
					Else
						strActual = "No search results found"
					End If
					
					If Instr(strActual, "Results 1-10 of")  Then
						boolResult = True
					Else
						boolResult = False
					End If
					strTestStep = "Performing sitewide search on " & strURL
					strExpected = "Non-zero search results for the term: #Cancer#"
					ResultWriteToFile boolResult, intTestStep, strTestStep, strExpected, strActual
					'################# Check the sitewide search functionality ##############
					
					'#######################Checking the English version ##############
'					Set oMainNavLinks = .WebElement("class:=nav-search-bar gradient header","html tag:=DIV","innerhtml:=<div class=.*nvcgSlMainNav.*").ChildObjects(oLinksDefn)
'					ClickThroughMainLinks oMainNavLinks, strURL
'					
					'#######################Checking the Spanish version ##############
					Browser("micclass:=Browser","index:=0").Navigate strURL
					.Link("html tag:=A","name:=Español","index:=0").Click
					WaitForBrowserSync
'					Set oMainNavLinks = .WebElement("class:=genSiteMainNav genSiteMainNavSpanish","html tag:=UL").ChildObjects(oLinksDefn)
'					ClickThroughMainLinks oMainNavLinks, strURL & "/espanol"

					
									
					'#######################Checking the PDQ page version, date ##############
					Browser("micclass:=Browser","index:=0").Navigate strURL & "/types/childhood-cancers/late-effects-pdq"
					msgbox "please check the PDQ page version, date and the Patient/HP toggle" 
					 
					'#######################Checking the PDQ images ##############
					Browser("micclass:=Browser","index:=0").Navigate strURL & "/images/cdr/live/CDR466533-750.jpg"
					msgbox "Check the PDQ images"

					'#######################Checking the Glossifier functionality ##############
					Browser("micclass:=Browser","index:=0").Navigate strURL & "/Common/PopUps/popDefinition.aspx?id=270740&version=Patient&language=English"
					msgbox "Check the the Glossifier pop-up"
					
					'#######################Checking the Blog series page, dynamic list contents, images ##############
					Browser("micclass:=Browser","index:=0").Navigate strURL & "/news-events/cancer-currents-blog"
					msgbox "Check for the presence of blog posts with thumbnail images on blog series page"
					msgbox "Check if Continue Reading links works on the blog series page"
					
					'#######################Checking the press release page, dynamic list contents,images ##############
					Browser("micclass:=Browser","index:=0").Navigate strURL & "/news-events"
					msgbox "Check for the presence of blog posts with thumbnail images on press release page"
					
					'#######################Checking the Exit Disclaimer ##############
					Browser("micclass:=Browser","index:=0").Navigate strURL & "/research/areas/genomics"
					WaitForBrowserSync
				     With Browser("micclass:=Browser","index:=0").Page("micclass:=page","index:=0")
	    				boolExists=.Link("html tag:=A","text:=Exit Disclaimer").Exist
							If boolExists=False Then
							Msgbox "Exit Disclaimer does not exist on CancerGov"
   					    	Else
   						 	.Link("html tag:=A","text:=Exit Disclaimer").Click
  					    	End If
						boolTemp=.webElement("html tag:=SPAN","innertext:=Website Linking Policy").Exist
   						 	If boolTemp=False Then
  			     			Msgbox "Did not land to the proper Exit Disclaimer page"
    					 	Else
							.webElement("html tag:=SPAN","innertext:=Website Linking Policy").Click
					    	End If
					End With
	                
	                
	                '####################### End of Checking the Exit Disclaimer ##############
					
					'RATS don't run on preview
					If Instr(ArrSites(idx), "preview") = 0 Then
						'RunAction "Action1 [BCRAT]", oneIteration, strURL
						
                         RunAction "Action1 [BCRAT1]", oneIteration,strURL
						 
						'RunAction "Action1 [BCRAT_ExitONError]", oneIteration, strURL
						 

						'RunAction "Action1 [CCRAT]", oneIteration, strURL
						
						 RunAction "Action1 [CCRAT1]", oneIteration, strURL
						
                        ' RunAction "Action1 [CCRAT_ExitOnError]", oneIteration, strURL
						
						'RunAction "Action1 [MRAT]", oneIteration, strURL
						
						 RunAction "Action1 [MRAT1]", oneIteration,strURL
						
						 'RunAction "Action1 [MRAT_ExitOnError]", oneIteration, strURL
						
					End If
				End if
			
			
			Case "cancergenome", "cancergenomepreview"
				If Len(CheckForErrorOnPage) = 0  Then
					'################# Check the sitewide search functionality ENGLISH for TCGA ##############
					Wait 2



					Browser("micclass:=browser").Page("micclass:=page").WebEdit("class:=main-search","title:=Search TCGA site","name:=swKeywordQuery","html tag:=INPUT").Set "Cancer"
					.Image("html tag:=INPUT","name:=Image","image type:=Image Button").Click
					WaitForBrowserSync
					BoolExists = .WebElement("html tag:=P","innerhtml:=Result.*","visible:=True","index:=0").Exist(0)
					If BoolExists = True Then
						strActual = .WebElement("html tag:=P","innerhtml:=Result.*","visible:=True","index:=0").GetROProperty("innertext")
					Else
						strActual = "No search results found"
					End If
					
					If Instr(strActual, "Results 1-10 of")  Then
						boolResult = True
					Else
						boolResult = False
					End If
					strTestStep = "Performing sitewide search on " & strURL
					strExpected = "Non-zero search results for the term: #Cancer#"
					ResultWriteToFile boolResult, intTestStep, strTestStep, strExpected, strActual
					intTestStep = intTestStep + 1
					
					'################# Check through the Main Links functionality for TCGA ##############
					
					oLinksDefn("visible").value = True
					Set oMainNavLinks = .WebElement("html tag:=UL","class:=level1","outertext:=HomeAbout Cancer GenomicsWhat is Cancer Genomics.*").ChildObjects(oLinksDefn)
					ClickThroughMainLinks oMainNavLinks, strURL
					
										
					Browser("micclass:=Browser","index:=0").Navigate strURL & "/cancersselected"
					FuncTCGA_GlossifiedTerms
					
				End If
				
				
				
				    '################# Check the Exit Disclaimer for TCGA ##############
				    
			Browser("micclass:=Browser","index:=0").Navigate strURL & "/newsevents/inthenews"
			WaitForBrowserSync
				
		       boolExists=Browser("micclass:=Browser").Page("micclass:=Page").Image("alt:=Exit Disclaimer","html tag:=IMG","image type:=Image Link","index:=0").Exist
		        If boolExists=False Then
		       	  Msgbox "The Exit Disclaimer link does not exist"
		        Else
		       	Browser("micclass:=Browser").Page("micclass:=Page").Image("alt:=Exit Disclaimer","html tag:=IMG","image type:=Image Link","index:=0").Click
		       	End If
		       	
		       	boolTemp=Browser("micclass:=Browser").Page("micclass:=Page").webElement("html tag:=SPAN","innertext:=Website Linking Policy").Exist
		       	If boolTemp=False Then
		       		Msgbox "The Website Linking Policy page did not load"
		       	Else
		       	    Browser("micclass:=Browser").Page("micclass:=Page").webElement("html tag:=SPAN","innertext:=Website Linking Policy").Click
		       	
		       	End If
		       	
		       	
		       	   
		       
				    
			Case "dceg","imaging", "dcegpreview","imagingpreview"
				If ArrSites(idx) = "dceg" Then
					Browser("micclass:=browser", "index:=0").Navigate strURL & "/news-events/linkage-newsletter"
					Msgbox "Check the DCEG Linkage newsletter subscription"
				End If
				If Len(CheckForErrorOnPage) = 0  Then
					If Not ArrSites(idx) = "ccop" Then
						
						
						'################# Check the sitewide search functionality DCEG ##############
						.WebEdit("default value:= Search","html id:=swKeywordQuery","html tag:=INPUT").Set "Cancer"
						.WebButton("html id:=swSearchButton","html tag:=INPUT","name:=Search","visible:=True").Click
						WaitForBrowserSync
						BoolExists = .WebElement("class:=genSiteSearchResultsCount","html tag:=P").Exist(0)
						If BoolExists = True Then
							strActual = .WebElement("class:=genSiteSearchResultsCount","html tag:=P").GetROProperty("innertext")
						Else
							strActual = "No search results found"
						End If
						
						If Instr(strActual, "Results 1-10 of")  Then
							boolResult = True
						Else
							boolResult = False
						End If
						strTestStep = "Performing sitewide search on " & strURL
						strExpected = "Non-zero search results for the term: #Cancer#"
						ResultWriteToFile boolResult, intTestStep, strTestStep, strExpected, strActual
						intTestStep = intTestStep + 1
						
						
					End If

					'################# Check the sitewide search functionality DCEG ##############
					oLinksDefn("visible").value = True
					Set oMainNavLinks = .WebElement("html id:=genSlotMainNav","html tag:=DIV","class:=clearFix").ChildObjects(oLinksDefn)
					ClickThroughMainLinks oMainNavLinks, strURL
				End If
				
				'################# Check the Exit Disclaimer DCEG ##############
				
				Browser("micclass:=Browser","index:=0").Navigate strURL & "/about/contact-dceg"
			    WaitForBrowserSync
			    
			    boolExists=Browser("micclass:=Browser").Page("micclass:=Page").Image("alt:=Exit Disclaimer","html tag:=IMG","image type:=Image Link","index:=1").Exist
			    If boolExists=False Then
			    	Msgbox "The Exit Disclaimer link does not exist"
			    Else
			        Browser("micclass:=Browser").Page("micclass:=Page").Image("alt:=Exit Disclaimer","html tag:=IMG","image type:=Image Link","index:=1").Click
			    
			    End If
			    
			    boolTemp=Browser("micclass:=Browser").Page("micclass:=Page").webElement("html tag:=SPAN","innertext:=Website Linking Policy").Exist
			    If boolTemp=False Then
			    	Msgbox "The Website Linking policy page is not loaded"
			    Else
			        Browser("micclass:=Browser").Page("micclass:=Page").webElement("html tag:=SPAN","innertext:=Website Linking Policy").Click
			    
			    End If
			    
			    

			Case "proteomics","proteomicspreview"
				If Len(CheckForErrorOnPage) = 0  Then
				
					'################# Check the sitewide search functionality Proteomics ##############
					Msgbox "Please go to IE-> Tools -> Comaptibility View Settings and UNCHECK Display intranet sites in compatibility view"
					
					.WebEdit("default value:= Search","html id:=swKeywordQuery","html tag:=INPUT").Set "Cancer"
					.WebButton("html id:=swSearchButton","html tag:=INPUT","name:=Search","visible:=True").Click
					WaitForBrowserSync
					BoolExists = .WebElement("class:=genSiteSearchResultsCount","html tag:=P").Exist(0)
					If BoolExists = True Then
						strActual = .WebElement("class:=genSiteSearchResultsCount","html tag:=P").GetROProperty("innertext")
					Else
						strActual = "No search results found"
					End If
					
					If Instr(strActual, "Results 1-10 of")  Then
						boolResult = True
					Else
						boolResult = False
					End If
					strTestStep = "Performing sitewide search on " & strURL
					strExpected = "Non-zero search results for the term: #Cancer#"
					ResultWriteToFile boolResult, intTestStep, strTestStep, strExpected, strActual
					intTestStep = intTestStep + 1
					
					'################# Check the sitewide search functionality Proteomics ##############
					
					
					oLinksDefn("visible").value = True
					oLinksDefn("url").value = "http.*://proteomics.*"
					Set oMainNavLinks = .WebElement("html id:=genSlotMainNav","html tag:=DIV","class:=clearFix").ChildObjects(oLinksDefn)
					ClickThroughMainLinks oMainNavLinks, strURL
				End If
				
			
				
				'################# Check the Exit Disclaimer Proteomics ##############
				Browser("micclass:=Browser","index:=0").Navigate strURL & "/resources/opendatapolicy"
			    WaitForBrowserSync
			    
			   ' Msgbox "Please go to IE-> Tools -> Comaptibility View Settings and uncheck Display intranet sites in compatibility view"
			    
			    wait 1
			    
			    boolExists=Browser("micclass:=Browser").Page("micclass:=Page").Image("alt:=Exit Disclaimer","html tag:=IMG","image type:=Image Link","index:=0").Exist
			    If boolExists=False Then
			       Msgbox "The Exit Disclaimer does not exist"
			     Else  
			       Browser("micclass:=Browser").Page("micclass:=Page").Image("alt:=Exit Disclaimer","html tag:=IMG","image type:=Image Link","index:=0").Click
			     End if
			     
			     boolTemp=Browser("micclass:=Browser").Page("micclass:=Page").webElement("html tag:=SPAN","innertext:=Website Linking Policy").Exist
			     
			     If boolExists=False Then
			     	Msgbox "The website linking policy page is not loaded"
			     Else
			        Browser("micclass:=Browser").Page("micclass:=Page").webElement("html tag:=SPAN","innertext:=Website Linking Policy").Click
			     
			     End If
			     
			     Msgbox "Please go to IE-> Tools -> Comaptibility View Settings and CHECK Display intranet sites in compatibility view"
			     
			End Select
	End With
	Set oLinksDefn = Nothing
'	Msgbox "Check the header and footer for " & strURL
Next




Function ClickThroughMainLinks (oMainNavLinks, oHomeLink)
	Dim ArrLinks(20)
	Dim ArrURLs(20)
	Tabs = oMainNavLinks.Count
	If Tabs > 0 Then
		For i=0 To Tabs-1
			ArrLinks(i) = Replace(oMainNavLinks(i).GetROProperty("innertext"), "?", "\?")
			ArrURLs(i) = oMainNavLinks(i).GetROProperty("url")
		Next
	End If
	'Click through each link
	With Browser("micclass:=browser").Page("micclass:=page")
		For i=1 To Tabs-1
'			.Link("html tag:=A","background color:=transparent", "index:=0","name:="& ArrLinks(i)).Click
			Browser("micclass:=Browser","index:=0").Navigate ArrURLs(i) 
			WaitForBrowserSync
			'####################### Page check and result logs #############################
			strActual = CheckForErrorOnPage
			If Len(strActual) > 0 Then 
				boolResult = False
'				Msgbox CheckForErrorOnPage  
			Else
				boolResult = True
				strActual = "Successful"
			End If
			strTestStep = "Checking the tab: " & ArrLinks(i) & " @ " & ArrURLs(i)
			strExpected = "Page loads successfully"
		
			ResultWriteToFile boolResult, intTestStep, strTestStep, strExpected, strActual
			intTestStep = intTestStep + 1
			'####################### Page check and result logs #############################
	
'			If Len(CheckForErrorOnPage) > 0 Then Msgbox CheckForErrorOnPage & " on the tab: " & ArrLinks(i)  End If
				
''				If Not oHomeLink Is Nothing Then 
'				If oHomeLink <> "" Then
'					Browser("micclass:=Browser","index:=0").Navigate oHomeLink 
'					WaitForBrowserSync
'				End If
		Next
	End With
	For i=0 To Tabs-1
		ArrLinks(i) = ""
	Next
	Set oMainNavLinks = Nothing
End Function
