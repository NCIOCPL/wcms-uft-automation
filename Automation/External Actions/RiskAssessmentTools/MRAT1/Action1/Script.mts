Dim strURL
Dim strEnvironment 
Dim RiskPercentage

If parameter("URL") <> "" Then
	strURL = parameter("URL")
Else
	strURL = "www-stage.cancer.gov" '"www.ocdev14.ha2.cancer.gov/"
End If

'######################################### Scenario 1: Mobile Male ######################################
Browser("micclass:=Browser","index:=0").Page("micclass:=page").Sync
Browser("micclass:=Browser","index:=0").Navigate strURL & "/melanomamobile/"

Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment_2").WebButton("Begin Risk Calculation").Click

With Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment_3")
	.WebList("ctl00$MainContent$ctl00$answer").Select "North"
	.WebButton("Next").Click @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebButton("Next")_;_script infofile_;_ZIP::ssf25.xml_;_
	.WebList("ctl00$MainContent$ctl00$answer").Select "Male" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebList("ctl00$MainContent$ctl00$answer")_;_script infofile_;_ZIP::ssf26.xml_;_
	.WebButton("Next").Click @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebButton("Next")_;_script infofile_;_ZIP::ssf27.xml_;_
	.WebRadioGroup("ctl00$MainContent$ctl00$answer").Select "1" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebRadioGroup("ctl00$MainContent$ctl00$answer")_;_script infofile_;_ZIP::ssf28.xml_;_
	.WebButton("Next").Click @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebButton("Next")_;_script infofile_;_ZIP::ssf29.xml_;_
	.WebList("ctl00$MainContent$ctl00$answer").Select "20" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebList("ctl00$MainContent$ctl00$answer")_;_script infofile_;_ZIP::ssf30.xml_;_
	.WebButton("Next").Click @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebButton("Next")_;_script infofile_;_ZIP::ssf31.xml_;_
	.WebList("ctl00$MainContent$ctl00$answer").Select "Yes"
	.WebButton("Next").Click
	.WebList("ctl00$MainContent$ctl00$answer").Select "Light" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebList("ctl00$MainContent$ctl00$answer")_;_script infofile_;_ZIP::ssf32.xml_;_
	.WebButton("Next").Click @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebButton("Next")_;_script infofile_;_ZIP::ssf33.xml_;_
	.WebList("ctl00$MainContent$ctl00$answer").Select "Less than two" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebList("ctl00$MainContent$ctl00$answer")_;_script infofile_;_ZIP::ssf36.xml_;_
	.WebButton("Next").Click @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebButton("Next")_;_script infofile_;_ZIP::ssf37.xml_;_
	.WebList("ctl00$MainContent$ctl00$answer").Select "Seven to sixteen"
	.WebButton("Next").Click
	.WebList("ctl00$MainContent$ctl00$answer").Select "Absent" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebList("ctl00$MainContent$ctl00$answer")_;_script infofile_;_ZIP::ssf38.xml_;_
	.WebButton("Next").Click
	.WebList("ctl00$MainContent$ctl00$answer").Select "Yes"
	.WebButton("Calculate Risk").Click
 End With
 
RiskPercentage = Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment_3").WebElement("Result").GetROProperty("innertext")
If RiskPercentage <> "0.04%" Then
	Msgbox "The Five-Year Absolute Risk of Melanoma: Mismatch" 
End If

'######################################### Scenario 2: Mobile Female ######################################

Browser("micclass:=Browser").Page("micclass:=page").Sync
Browser("micclass:=Browser").Navigate strURL & "/melanomamobile/"

Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment_2").WebButton("Begin Risk Calculation").Click

With Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment_3")
	.WebList("ctl00$MainContent$ctl00$answer").Select "North"
	.WebButton("Next").Click @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebButton("Next")_;_script infofile_;_ZIP::ssf25.xml_;_
	.WebList("ctl00$MainContent$ctl00$answer").Select "Female" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebList("ctl00$MainContent$ctl00$answer")_;_script infofile_;_ZIP::ssf26.xml_;_
	.WebButton("Next").Click @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebButton("Next")_;_script infofile_;_ZIP::ssf27.xml_;_
	.WebRadioGroup("ctl00$MainContent$ctl00$answer").Select "1" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebRadioGroup("ctl00$MainContent$ctl00$answer")_;_script infofile_;_ZIP::ssf28.xml_;_
	.WebButton("Next").Click @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebButton("Next")_;_script infofile_;_ZIP::ssf29.xml_;_
	.WebList("ctl00$MainContent$ctl00$answer").Select "20" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebList("ctl00$MainContent$ctl00$answer")_;_script infofile_;_ZIP::ssf30.xml_;_
	.WebButton("Next").Click @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebButton("Next")_;_script infofile_;_ZIP::ssf31.xml_;_
	.WebList("ctl00$MainContent$ctl00$answer").Select "Light" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebList("ctl00$MainContent$ctl00$answer")_;_script infofile_;_ZIP::ssf32.xml_;_
	.WebButton("Next").Click @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebButton("Next")_;_script infofile_;_ZIP::ssf33.xml_;_
	.WebList("ctl00$MainContent$ctl00$answer").Select "Moderately tanned" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebList("ctl00$MainContent$ctl00$answer")_;_script infofile_;_ZIP::ssf34.xml_;_
	.WebButton("Next").Click @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebButton("Next")_;_script infofile_;_ZIP::ssf35.xml_;_
	.WebList("ctl00$MainContent$ctl00$answer").Select "Less than five" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebList("ctl00$MainContent$ctl00$answer")_;_script infofile_;_ZIP::ssf36.xml_;_
	.WebButton("Next").Click @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebButton("Next")_;_script infofile_;_ZIP::ssf37.xml_;_
	.WebList("ctl00$MainContent$ctl00$answer").Select "Absent" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 3").WebList("ctl00$MainContent$ctl00$answer")_;_script infofile_;_ZIP::ssf38.xml_;_
	.WebButton("Calculate Risk").Click
End With @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment").Image("Calculate Risk")_;_script infofile_;_ZIP::ssf22.xml_;_
 
Wait 2
RiskPercentage = Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment_3").WebElement("Result").GetROProperty("innertext")
If RiskPercentage <> "0.01%" Then
	Msgbox "The Five-Year Absolute Risk of Melanoma: Mismatch" 
End If
'
'######################################### Scenario 3: Female ######################################

Browser("micclass:=Browser").Page("micclass:=page").Sync
Browser("micclass:=Browser").Navigate strURL & "/melanomarisktool/"

 With Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment")
 	.WebList("region").Select "North"
	.WebList("sex").Select "Female" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 2").WebList("sex")_;_script infofile_;_ZIP::ssf3.xml_;_
	.WebList("race").Select "Non-Hispanic White" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 2").WebList("race")_;_script infofile_;_ZIP::ssf4.xml_;_
	.WebList("age").Select "20" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 2").WebList("sunburn")_;_script infofile_;_ZIP::ssf6.xml_;_
	.WebList("complexion").Select "Light" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 2").WebList("complexion")_;_script infofile_;_ZIP::ssf7.xml_;_
	.WebList("tanning").Select "Moderately tanned"
	.WebList("small_moles_females").Select "Less than five"
	.WebList("freckling").Select "Absent" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 2").WebList("solar damage")_;_script infofile_;_ZIP::ssf11.xml_;_
	.Image("Calculate Risk").Click
 End With @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment").Image("Calculate Risk")_;_script infofile_;_ZIP::ssf22.xml_;_
Wait 2
RiskPercentage = Browser("Melanoma Risk Assessment").Page("Results - MRAT").WebElement("ResultPercentage").GetROProperty("innertext")
If RiskPercentage <> "0.01%" Then
	Msgbox "The Five-Year Absolute Risk of Melanoma: Mismatch" 
End If


'######################################### Scenario 4: Male ######################################
Browser("micclass:=Browser").Page("micclass:=page").Sync
Browser("micclass:=Browser").Navigate strURL & "/melanomarisktool/"

 With Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment")
 	.WebList("region").Select "North"
	.WebList("sex").Select "Male" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 2").WebList("sex")_;_script infofile_;_ZIP::ssf3.xml_;_
	.WebList("race").Select "Non-Hispanic White" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 2").WebList("race")_;_script infofile_;_ZIP::ssf4.xml_;_
	.WebList("age").Select "20" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 2").WebList("age")_;_script infofile_;_ZIP::ssf5.xml_;_
	.WebList("sunburn").Select "Yes" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 2").WebList("sunburn")_;_script infofile_;_ZIP::ssf6.xml_;_
	.WebList("complexion").Select "Light" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 2").WebList("complexion")_;_script infofile_;_ZIP::ssf7.xml_;_
	.WebList("large_moles").Select "Less than two" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 2").WebList("large moles")_;_script infofile_;_ZIP::ssf8.xml_;_
	.WebList("small_moles_males").Select "Seven to sixteen" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 2").WebList("small moles males")_;_script infofile_;_ZIP::ssf9.xml_;_
	.WebList("freckling").Select "Absent" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 2").WebList("freckling")_;_script infofile_;_ZIP::ssf10.xml_;_
	.WebList("solar_damage").Select "Yes" @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 2").WebList("solar damage")_;_script infofile_;_ZIP::ssf11.xml_;_
	.Image("Calculate Risk").Click
 End With
 @@ hightlight id_;_Browser("Melanoma Risk Assessment").Page("Melanoma Risk Assessment 2").Image("Calculate Risk")_;_script infofile_;_ZIP::ssf12.xml_;_
RiskPercentage = Browser("Melanoma Risk Assessment").Page("Results - MRAT").WebElement("ResultPercentage").GetROProperty("innertext")
If RiskPercentage <> "0.04%" Then
	Msgbox "The Five-Year Absolute Risk of Melanoma: Mismatch" 
End If
