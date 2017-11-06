Dim strURL
Dim strSite
Dim strFiveYrRisk, strTenYrRisk, strLifetimeYrRisk

If parameter("URL") <> "" Then
	strSite = parameter("URL")
Else
	strSite = "www.cancer.gov" 
End If

strURL = strSite & "/colorectalcancerrisk/"

Browser("micclass:=browser", "index:=0").Navigate strURL 
Browser("Colorectal Cancer Risk").Page("Colorectal Cancer Risk").Link("Risk Calculator >").Click

'###################################### Scenario1 - Female ####################################################
With Browser("Colorectal Cancer Risk_2").Page("Colorectal Cancer Risk")
	.WebRadioGroup("ctl00$cphMain$WUCDemo$rblhispa").Select "No"
	.WebRadioGroup("ctl00$cphMain$WUCDemo$rblRace").Select "White" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebRadioGroup("ctl00$cphMain$WUCDemo$rblRace")_;_script infofile_;_ZIP::ssf2.xml_;_
	.WebList("ctl00$cphMain$WUCDemo$ddlCurre").Select "60" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebList("ctl00$cphMain$WUCDemo$ddlCurre")_;_script infofile_;_ZIP::ssf3.xml_;_
	.WebRadioGroup("ctl00$cphMain$WUCDemo$rblSex").Select "Female" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebRadioGroup("ctl00$cphMain$WUCDemo$rblSex")_;_script infofile_;_ZIP::ssf4.xml_;_
	.WebEdit("ctl00$cphMain$WUCDemo$txtFeet").Set "5" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebEdit("ctl00$cphMain$WUCDemo$txtFeet")_;_script infofile_;_ZIP::ssf5.xml_;_
	.WebEdit("ctl00$cphMain$WUCDemo$txtInch").Set "5" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebEdit("ctl00$cphMain$WUCDemo$txtInch")_;_script infofile_;_ZIP::ssf6.xml_;_
	.WebEdit("ctl00$cphMain$WUCDemo$txtWeigh").Set "150" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebEdit("ctl00$cphMain$WUCDemo$txtWeigh")_;_script infofile_;_ZIP::ssf7.xml_;_
	.Image("Next").Click @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").Image("Next")_;_script infofile_;_ZIP::ssf8.xml_;_
	.WebList("ctl00$cphMain$WUCDiet$ddlVeggi").Select "3-4 servings per week" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebList("ctl00$cphMain$WUCDiet$ddlVeggi")_;_script infofile_;_ZIP::ssf9.xml_;_
	.WebList("ctl00$cphMain$WUCDiet$ddlVeggi_2").Select "Between 3 cups and 5 cups" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebList("ctl00$cphMain$WUCDiet$ddlVeggi 2")_;_script infofile_;_ZIP::ssf10.xml_;_
	.Image("Next").Click 39,11 @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").Image("Next")_;_script infofile_;_ZIP::ssf11.xml_;_
	.WebRadioGroup("ctl00$cphMain$WUCMedicalHistor").Select "Yes" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebRadioGroup("ctl00$cphMain$WUCMedicalHistor")_;_script infofile_;_ZIP::ssf12.xml_;_
	.WebRadioGroup("ctl00$cphMain$WUCMedicalHistor_2").Select "No" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebRadioGroup("ctl00$cphMain$WUCMedicalHistor 2")_;_script infofile_;_ZIP::ssf13.xml_;_
	.Image("Next").Click 37,12 @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").Image("Next")_;_script infofile_;_ZIP::ssf14.xml_;_
	.WebRadioGroup("ctl00$cphMain$WUCMedication$rb").Select "Yes" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebRadioGroup("ctl00$cphMain$WUCMedication$rb")_;_script infofile_;_ZIP::ssf15.xml_;_
	.WebRadioGroup("ctl00$cphMain$WUCMedication$rb_2").Select "Yes" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebRadioGroup("ctl00$cphMain$WUCMedication$rb 2")_;_script infofile_;_ZIP::ssf16.xml_;_
	.Image("Next").Click 24,11 @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").Image("Next")_;_script infofile_;_ZIP::ssf17.xml_;_
 @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").Image("Next")_;_script infofile_;_ZIP::ssf3.xml_;_
 @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").Image("Calculate")_;_script infofile_;_ZIP::ssf4.xml_;_
	
	.WebList("ctl00$cphMain$WUCPhysicalActiv").Select "1" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebList("ctl00$cphMain$WUCPhysicalActiv")_;_script infofile_;_ZIP::ssf18.xml_;_
	.WebList("ctl00$cphMain$WUCPhysicalActiv_2").Select "Up to 1 hour per week" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebList("ctl00$cphMain$WUCPhysicalActiv 2")_;_script infofile_;_ZIP::ssf19.xml_;_
	.WebList("ctl00$cphMain$WUCPhysicalActiv_3").Select "12" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebList("ctl00$cphMain$WUCPhysicalActiv 3")_;_script infofile_;_ZIP::ssf20.xml_;_
	.WebList("ctl00$cphMain$WUCPhysicalActiv_4").Select "More than 4 hours per week" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebList("ctl00$cphMain$WUCPhysicalActiv 4")_;_script infofile_;_ZIP::ssf21.xml_;_
	.Image("Next").Click 31,9 @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").Image("Next")_;_script infofile_;_ZIP::ssf22.xml_;_
	.WebRadioGroup("ctl00$cphMain$WUCMiscWoman$rbl").Select "No" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebRadioGroup("ctl00$cphMain$WUCMiscWoman$rbl")_;_script infofile_;_ZIP::ssf23.xml_;_
	.WebList("ctl00$cphMain$WUCMiscWoman$ddl").Select "2 years ago or more" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebList("ctl00$cphMain$WUCMiscWoman$ddl")_;_script infofile_;_ZIP::ssf24.xml_;_
	.WebRadioGroup("ctl00$cphMain$WUCMiscWoman$rbl_2").Select "No" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebRadioGroup("ctl00$cphMain$WUCMiscWoman$rbl 2")_;_script infofile_;_ZIP::ssf25.xml_;_
	.Image("Next").Click 54,10 @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").Image("Next")_;_script infofile_;_ZIP::ssf26.xml_;_
	.WebRadioGroup("ctl00$cphMain$WUCFamily$rblRel").Select "No" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebRadioGroup("ctl00$cphMain$WUCFamily$rblRel")_;_script infofile_;_ZIP::ssf27.xml_;_
	Browser("Colorectal Cancer Risk_2").Page("Colorectal Cancer Risk").Image("Calculate").Click 5,5
End With
 @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").Image("Calculate")_;_script infofile_;_ZIP::ssf28.xml_;_
Wait 4
Browser("Colorectal Cancer Risk_2").Page("Colorectal Cancer Risk_2").Link("5 year riskAverage").Click
strFiveYrRisk = Browser("Colorectal Cancer Risk_2").Page("Colorectal Cancer Risk_2").Link("5 year riskAverage").GetROProperty("innertext") @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk 2").Link("5 year riskAverage:0.5%You:")_;_script infofile_;_ZIP::ssf29.xml_;_
Browser("Colorectal Cancer Risk_2").Page("Colorectal Cancer Risk_2").Link("10 year riskAverage").Click
strTenYrRisk = Browser("Colorectal Cancer Risk_2").Page("Colorectal Cancer Risk_2").Link("10 year riskAverage").GetROProperty("innertext") @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk 2").Link("10 year riskAverage:1.1%You:")_;_script infofile_;_ZIP::ssf30.xml_;_
Browser("Colorectal Cancer Risk_2").Page("Colorectal Cancer Risk_2").Link("Lifetime riskAverage").Click
strLifetimeYrRisk = Browser("Colorectal Cancer Risk_2").Page("Colorectal Cancer Risk_2").Link("Lifetime riskAverage").GetROProperty("innertext") @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk 2").Link("Lifetime riskAverage:4.9%You:")_;_script infofile_;_ZIP::ssf31.xml_;_

If strFiveYrRisk <> "5 year riskAverage:0.5%You: 0.2%" Then
	Msgbox "5 year riskAverage: Failed"
End If
If strTenYrRisk <> "10 year riskAverage:1.1%You: 0.4%" Then
	Msgbox "10 year riskAverage: Failed"
End If
If strLifetimeYrRisk <> "Lifetime riskAverage:4.9%You: 1.7%" Then
	Msgbox "Lifetime riskAverage: Failed"
End If
Browser("Colorectal Cancer Risk_2").CloseAllTabs
'###################################### Scenario 2 - Male ####################################################

Browser("micclass:=browser", "index:=0").Navigate strURL 
Browser("Colorectal Cancer Risk").Page("Colorectal Cancer Risk").Link("Risk Calculator >").Click

With Browser("Colorectal Cancer Risk_2").Page("Colorectal Cancer Risk")
	.WebRadioGroup("ctl00$cphMain$WUCDemo$rblhispa").Select "No"
	.WebRadioGroup("ctl00$cphMain$WUCDemo$rblRace").Select "White" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebRadioGroup("ctl00$cphMain$WUCDemo$rblRace")_;_script infofile_;_ZIP::ssf2.xml_;_
	.WebList("ctl00$cphMain$WUCDemo$ddlCurre").Select "60" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebList("ctl00$cphMain$WUCDemo$ddlCurre")_;_script infofile_;_ZIP::ssf3.xml_;_
	.WebRadioGroup("ctl00$cphMain$WUCDemo$rblSex").Select "Male" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebRadioGroup("ctl00$cphMain$WUCDemo$rblSex")_;_script infofile_;_ZIP::ssf4.xml_;_
	.WebEdit("ctl00$cphMain$WUCDemo$txtFeet").Set "5" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebEdit("ctl00$cphMain$WUCDemo$txtFeet")_;_script infofile_;_ZIP::ssf5.xml_;_
	.WebEdit("ctl00$cphMain$WUCDemo$txtInch").Set "5" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebEdit("ctl00$cphMain$WUCDemo$txtInch")_;_script infofile_;_ZIP::ssf6.xml_;_
	.WebEdit("ctl00$cphMain$WUCDemo$txtWeigh").Set "150" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebEdit("ctl00$cphMain$WUCDemo$txtWeigh")_;_script infofile_;_ZIP::ssf7.xml_;_
	.Image("Next").Click 37,10 @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").Image("Next")_;_script infofile_;_ZIP::ssf8.xml_;_
	.WebList("ctl00$cphMain$WUCDiet$ddlVeggi").Select "3-4 servings per week" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebList("ctl00$cphMain$WUCDiet$ddlVeggi")_;_script infofile_;_ZIP::ssf9.xml_;_
	.WebList("ctl00$cphMain$WUCDiet$ddlVeggi_2").Select "Between 3 cups and 5 cups" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebList("ctl00$cphMain$WUCDiet$ddlVeggi 2")_;_script infofile_;_ZIP::ssf10.xml_;_
	.Image("Next").Click 39,11 @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").Image("Next")_;_script infofile_;_ZIP::ssf11.xml_;_
	.WebRadioGroup("ctl00$cphMain$WUCMedicalHistor").Select "Yes" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebRadioGroup("ctl00$cphMain$WUCMedicalHistor")_;_script infofile_;_ZIP::ssf12.xml_;_
	.WebRadioGroup("ctl00$cphMain$WUCMedicalHistor_2").Select "No" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebRadioGroup("ctl00$cphMain$WUCMedicalHistor 2")_;_script infofile_;_ZIP::ssf13.xml_;_
	.Image("Next").Click 37,12 @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").Image("Next")_;_script infofile_;_ZIP::ssf14.xml_;_
	.WebRadioGroup("ctl00$cphMain$WUCMedication$rb").Select "Yes" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebRadioGroup("ctl00$cphMain$WUCMedication$rb")_;_script infofile_;_ZIP::ssf15.xml_;_
	.WebRadioGroup("ctl00$cphMain$WUCMedication$rb_2").Select "Yes" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebRadioGroup("ctl00$cphMain$WUCMedication$rb 2")_;_script infofile_;_ZIP::ssf16.xml_;_
	.Image("Next").Click 24,11 @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").Image("Next")_;_script infofile_;_ZIP::ssf17.xml_;_
	.WebList("ctl00$cphMain$WUCPhysicalActiv").Select "1" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebList("ctl00$cphMain$WUCPhysicalActiv")_;_script infofile_;_ZIP::ssf18.xml_;_
	.WebList("ctl00$cphMain$WUCPhysicalActiv_2").Select "Up to 1 hour per week" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebList("ctl00$cphMain$WUCPhysicalActiv 2")_;_script infofile_;_ZIP::ssf19.xml_;_
	.WebList("ctl00$cphMain$WUCPhysicalActiv_3").Select "12" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebList("ctl00$cphMain$WUCPhysicalActiv 3")_;_script infofile_;_ZIP::ssf20.xml_;_
	.WebList("ctl00$cphMain$WUCPhysicalActiv_4").Select "More than 4 hours per week" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebList("ctl00$cphMain$WUCPhysicalActiv 4")_;_script infofile_;_ZIP::ssf21.xml_;_
	.Image("Next").Click 31,9 @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebRadioGroup("ctl00$cphMain$WUCMiscWoman$rbl")_;_script infofile_;_ZIP::ssf23.xml_;_
	
	.WebRadioGroup("ctl00$cphMain$WUCMiscellaneous").Select "Yes" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebRadioGroup("ctl00$cphMain$WUCMiscellaneous")_;_script infofile_;_ZIP::ssf32.xml_;_
	.WebList("ctl00$cphMain$WUCMiscellaneous").Select "20" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebList("ctl00$cphMain$WUCMiscellaneous")_;_script infofile_;_ZIP::ssf33.xml_;_
	.WebRadioGroup("ctl00$cphMain$WUCMiscellaneous_2").Select "Yes" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebRadioGroup("ctl00$cphMain$WUCMiscellaneous 2")_;_script infofile_;_ZIP::ssf34.xml_;_
	.WebList("ctl00$cphMain$WUCMiscellaneous_2").Select "1 to 10 cigarettes a day" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebList("ctl00$cphMain$WUCMiscellaneous 2")_;_script infofile_;_ZIP::ssf35.xml_;_
	.Image("Next").Click 43,16 @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").Image("Next")_;_script infofile_;_ZIP::ssf36.xml_;_
	.WebRadioGroup("ctl00$cphMain$WUCFamily$rblRel").Select "Yes" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebRadioGroup("ctl00$cphMain$WUCFamily$rblRel")_;_script infofile_;_ZIP::ssf37.xml_;_
	.WebRadioGroup("ctl00$cphMain$WUCFamily$rblRel_2").Select "One" @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").WebRadioGroup("ctl00$cphMain$WUCFamily$rblRel 2")_;_script infofile_;_ZIP::ssf38.xml_;_
	Browser("Colorectal Cancer Risk_2").Page("Colorectal Cancer Risk").Image("Calculate").Click 5,5 @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk 2").Link("Lifetime riskAverage:5.7%You:")_;_script infofile_;_ZIP::ssf42.xml_;_
	.Sync
End With
 @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk").Image("Calculate")_;_script infofile_;_ZIP::ssf28.xml_;_
Wait 4
strFiveYrRisk = Browser("Colorectal Cancer Risk_2").Page("Colorectal Cancer Risk_2").Link("5 year riskAverage").GetROProperty("innertext") @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk 2").Link("5 year riskAverage:0.5%You:")_;_script infofile_;_ZIP::ssf29.xml_;_
strTenYrRisk = Browser("Colorectal Cancer Risk_2").Page("Colorectal Cancer Risk_2").Link("10 year riskAverage").GetROProperty("innertext") @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk 2").Link("10 year riskAverage:1.1%You:")_;_script infofile_;_ZIP::ssf30.xml_;_
strLifetimeYrRisk = Browser("Colorectal Cancer Risk_2").Page("Colorectal Cancer Risk_2").Link("Lifetime riskAverage").GetROProperty("innertext") @@ hightlight id_;_Browser("Colorectal Cancer Risk 2").Page("Colorectal Cancer Risk 2").Link("Lifetime riskAverage:4.9%You:")_;_script infofile_;_ZIP::ssf31.xml_;_

If strFiveYrRisk <> "5 year riskAverage:0.7%You: 0.4%" Then
	Msgbox "5 year riskAverage: Failed"
End If
If strTenYrRisk <> "10 year riskAverage:1.6%You: 1.1%" Then
	Msgbox "10 year riskAverage: Failed"
End If
If strLifetimeYrRisk <> "Lifetime riskAverage:5.7%You: 3.5%" Then
	Msgbox "Lifetime riskAverage: Failed"
End If

Browser("Colorectal Cancer Risk_2").CloseAllTabs
