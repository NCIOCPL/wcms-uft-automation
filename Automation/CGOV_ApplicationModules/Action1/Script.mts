'Set oBrowserPage = Browser("micclass:=Browser","index:=0").Page("micclass:=page","index:=0")
'Set oLink = Description.Create
'oLink("micclass").value = "Link"
'oLink("html tag").value =  "A"
'oLink("class").value =  "protocol-abstract-link"
'
'Set Ocnt23 = oBrowserPage.ChildObjects(oLink)
'tmpcount =  Ocnt23.count
'msgbox tmpcount & " CT Search Results are being displayed"
'If tmpcount <> 25 Then
'	Msgbox "System is not displaying 25 results for the default advanced search"
'End If
''Click the first link
'Ocnt23(0).click


'Browser("micclass:=Browser","index:=0").Page("micclass:=page","index:=0").Link("class:=protocol-abstract-link","html tag:=A","index:=0").click



   strEnv = Inputbox ("Please enter one of these Environments: www1, dev13, dev14, blue red pink qa dt stage prod")
	boolRunAppModuleTests = True
	If strEnv = "" Then strEnv = "stage" End If
	
	
	If strEnv="prod" Then
		strMode = Inputbox ("Please enter 1 for LIVE, 2 for PREVIEW")
         
           If strMode = "" Then
	          strMode = 1
            End If
            
          If strMode=1 Then
           	strEnv="prod"
           	Else 
           	strEnv="preview"
           End If
           
	End If
	
	
	Environment.Value("ResultsPath") = "L:\OCPL\ODDC\CPSB\QAStuff\Veena\Automation\Results\WCMS\ApModules"
	                                   
	Environment.Value("TestcaseFile") = "WCMS_AppModuleReport"
	Environment.Value("strTestName") = "WCMS AppModule Test"
	Public intTestStep	
	intTestStep = 1
	Set ParentBr = Browser("micclass:=browser").Page("micclass:=page")
	
	ArrURLs = Array("Search", "GeneticsDirectory", "DrugDictionary", "DictionaryOfCancerTerms", "diccionario", "ClinicalTrials")
	strTmp = SetSiteURL (strEnv)
	
'	 strURL = strTmp & "/about-cancer/treatment/clinical-trials/search"	'# Basic Clinical Trials Search #
'	 FuncBasicClinicalTrialsSearch strURL
'	
	
	strURL = strTmp & "/about-cancer/treatment/clinical-trials/advanced-search"	'# Advanced Clinical Trials Search #
	FuncClinicalTrialsSearch strURL
	
	strURL = strTmp & "/research/nci-role/cancer-centers/find"	'# Cancer Center Maps #
	FuncCancerCenterMaps strURL

	strURL = strTmp & "/about-cancer/causes-prevention/genetics/directory"
	FuncNCICancerGeneticsServicesDirectory strURL	'# NCI Cancer Genetics Services Directory: Search #

	strURL = strTmp
	FuncCancerGovSiteWideSearch strURL	'# Site Search #
	
	strURL = strTmp & "/syndication/widgets"   '# Dictionary Widget #
    FuncDictWidget strURL

	strURL = strTmp & "/espanol/publicaciones/diccionario" '# Diccionario de cáncer (Spanish NCI Dictionary of Cancer Terms) #
	FuncSpanishNCIDictionaryOfCancerTerms strURL

	strURL = strTmp & "/publications/dictionaries/genetics-dictionary" '# NCI Dictionary of Genetics Terms #
	FuncNCIDictionaryOfGeneticTerms strURL
	
	strURL = strTmp & "/publications/dictionaries/cancer-terms" '# NCI Dictionary of Cancer Terms #
	FuncEnglishNCIDictionaryOfCancerTerms strURL
	
	strURL = strTmp & "/publications/dictionaries/cancer-drug"   '# NCI Drug Dictionary #
	FuncNCIDrugDictionary strURL
	



	strURL = strTmp
	TestCTHP = Inputbox ("Do you want to check CTHP pages? Y/N")
	If Lcase(TestCTHP = "y") Then
		FuncCheckCTHP strURL
	End If
