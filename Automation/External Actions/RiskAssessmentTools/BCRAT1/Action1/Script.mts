Dim strURL
Dim strEnvironment 
Dim strResultFiveYr, strResultLifeTime

'strURL = Inputbox ("Please enter the complete url")
If parameter("URL") <> "" Then
	strURL = parameter("URL")
Else
	
	 strURL = "www.cancer.gov" 
End If



'################################# Scenario 2 Without children #####################################
Browser("micclass:=Browser").Page("micclass:=page").Sync
Browser("micclass:=Browser").Navigate strURL & "/bcrisktool/Default.aspx" 

With Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment")
	.WebList("history").Select "No"
	.WebList("genetics").Select "No"
	.WebList("current_age").Select "36" @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment").WebList("age at menarche")_;_script infofile_;_ZIP::ssf3.xml_;_
	.WebList("age_at_menarche").Select "7 to 11" @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment").WebList("age at menarche")_;_script infofile_;_ZIP::ssf4.xml_;_
	.WebList("age_at_first_live_birth").Select "No births" @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment").WebList("age at first live birth")_;_script infofile_;_ZIP::ssf5.xml_;_
	.WebList("related_with_breast_cancer").Select "0" @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment").WebList("related with breast cancer")_;_script infofile_;_ZIP::ssf6.xml_;_
	.WebList("ever_had_biopsy").Select "Yes" @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment").WebList("ever had biopsy")_;_script infofile_;_ZIP::ssf7.xml_;_
	.WebList("previous_biopsies").Select "1" @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment").WebList("previous biopsies")_;_script infofile_;_ZIP::ssf8.xml_;_
	.WebList("biopsy_with_hyperplasia").Select "No" @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment").WebList("biopsy with hyperplasia")_;_script infofile_;_ZIP::ssf9.xml_;_
	.WebList("race").Select "Asian-American" @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment").WebList("race")_;_script infofile_;_ZIP::ssf10.xml_;_
	.WebList("subrace").Select "Japanese" @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment").WebList("subrace")_;_script infofile_;_ZIP::ssf11.xml_;_
	.Image("Calculate Risk").Click
End With
 @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment").Image("Calculate Risk")_;_script infofile_;_ZIP::ssf12.xml_;_
 
 
 Wait 2
 
 strResultFiveYr = Browser("Breast Cancer Risk Assessment").Page("Results - BCRA").WebElement("FiveYrRisk").GetROProperty("innertext")
 strResultLifeTime = Browser("Breast Cancer Risk Assessment").Page("Results - BCRA").WebElement("LifeTimeRisk").GetROProperty("innertext")
 
 If strResultFiveYr <> "This woman (age 36): 0.5%Average woman (age 36): 0.3%" Then
 	Msgbox "5 year % Failed: Expected--- This woman (age 36): 0.5%Average woman (age 36): 0.3%"
 End If
  If strResultLifeTime <> "This woman (to age 90): 18.8%Average woman (to age 90): 12%" Then
 	Msgbox "Lifetime % Failed: Expected--- This woman (to age 90): 18.8%Average woman (to age 90): 12%"
 End If

'################################# Scenario 1 With children #####################################
Browser("micclass:=Browser").Page("micclass:=page").Sync
Browser("micclass:=Browser").Navigate strURL & "/bcrisktool/Default.aspx" 

With Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment")
	.WebList("history").Select "No"
	.WebList("genetics").Select "No" @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment 2").WebList("genetics")_;_script infofile_;_ZIP::ssf13.xml_;_
	.WebList("current_age").Select "36" @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment").WebList("age at menarche")_;_script infofile_;_ZIP::ssf3.xml_;_
	.WebList("age_at_menarche").Select "7 to 11" @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment").WebList("age at menarche")_;_script infofile_;_ZIP::ssf4.xml_;_
	.WebList("age_at_first_live_birth").Select "< 20" @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment").WebList("age at first live birth")_;_script infofile_;_ZIP::ssf5.xml_;_
	.WebList("related_with_breast_cancer").Select "0" @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment").WebList("related with breast cancer")_;_script infofile_;_ZIP::ssf6.xml_;_
	.WebList("ever_had_biopsy").Select "Yes" @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment").WebList("ever had biopsy")_;_script infofile_;_ZIP::ssf7.xml_;_
	.WebList("previous_biopsies").Select "1" @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment").WebList("previous biopsies")_;_script infofile_;_ZIP::ssf8.xml_;_
	.WebList("biopsy_with_hyperplasia").Select "No" @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment").WebList("biopsy with hyperplasia")_;_script infofile_;_ZIP::ssf9.xml_;_
	.WebList("race").Select "Asian-American" @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment").WebList("race")_;_script infofile_;_ZIP::ssf10.xml_;_
	.WebList("subrace").Select "Japanese" @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment").WebList("subrace")_;_script infofile_;_ZIP::ssf11.xml_;_
	.Image("Calculate Risk").Click
End With
 @@ hightlight id_;_Browser("Breast Cancer Risk Assessment").Page("Breast Cancer Risk Assessment").Image("Calculate Risk")_;_script infofile_;_ZIP::ssf12.xml_;_
 Wait 2
 
 strResultFiveYr = Browser("Breast Cancer Risk Assessment").Page("Results - BCRA").WebElement("FiveYrRisk").GetROProperty("innertext")
 strResultLifeTime = Browser("Breast Cancer Risk Assessment").Page("Results - BCRA").WebElement("LifeTimeRisk").GetROProperty("innertext")
 
 If strResultFiveYr <> "This woman (age 36): 0.3%Average woman (age 36): 0.3%" Then
 	Msgbox "5 year % Failed: Expected--- This woman (age 36) 0.3%Average woman (age 36): 0.3%"
 End If
  If strResultLifeTime <> "This woman (to age 90): 11.3%Average woman (to age 90): 12%" Then
 	Msgbox "Lifetime % Failed: Expected--- This woman (to age 90): 11.3%Average woman (to age 90): 12%"
 End If
