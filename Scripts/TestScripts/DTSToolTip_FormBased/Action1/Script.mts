Browser("Tooltip - GlobalSQA_2").Page("Home - GlobalSQA").Link("Tooltip").Click @@ script infofile_;_ZIP::ssf2.xml_;_
Browser("Tooltip - GlobalSQA_2").Page("Tooltip - GlobalSQA_2").WebElement("Forms Based").Click @@ script infofile_;_ZIP::ssf3.xml_;_


' Usage of the function in your script
SelfClosingMsgBox "Please provide your firstname", 2, "Tooltip Text"

' Store the tooltip text in a variable
tooltipText =Browser("Tooltip - GlobalSQA_2").Page("Tooltip - GlobalSQA").Frame("Frame").WebEdit("firstname").GetROProperty("title")

' Verify the tooltip text
Dim result
If tooltipText = "Please provide your firstname." Then
    result = "Passed"
Else
    result = "Failed"
End If

' Write result to Excel
WriteResultToExcel "DTSToolTip_FormBased", "Verify Tooltip of a firstname textbox", "Please provide your firstname.", tooltipText, result, "G:\UFT Projects\GlobalSqa_Automation_Project\Results\Reports\TestReport.xlsx"

 @@ script infofile_;_ZIP::ssf5.xml_;_
Browser("Tooltip - GlobalSQA_2").Page("Tooltip - GlobalSQA").Frame("Frame").WebEdit("firstname").Set "rytika" @@ script infofile_;_ZIP::ssf6.xml_;_


' Usage of the function in your script
SelfClosingMsgBox "Please provide your lastname", 2, "Tooltip Text"

' Store the tooltip text in a variable
tooltipText1 =Browser("Tooltip - GlobalSQA_2").Page("Tooltip - GlobalSQA").Frame("Frame").WebEdit("lastname").GetROProperty("title")

' Verify the tooltip text
Dim result1
If tooltipText1 = "Please provide also your lastname." Then
    result1 = "Passed"
Else
    result1 = "Failed"
End If

' Write result to Excel
WriteResultToExcel "DTSToolTip_FormBased", "Verify Tooltip of a lastname textbox", "Please provide also your lastname.", tooltipText1, result1, "G:\UFT Projects\GlobalSqa_Automation_Project\Results\Reports\TestReport.xlsx"

Browser("Tooltip - GlobalSQA_2").Page("Tooltip - GlobalSQA").Frame("Frame").WebEdit("lastname").Set "tuiyiu" @@ script infofile_;_ZIP::ssf7.xml_;_



' Usage of the function in your script
SelfClosingMsgBox "Your home or work address", 2, "Tooltip Text"

' Store the tooltip text in a variable
tooltipText2 =Browser("Tooltip - GlobalSQA_2").Page("Tooltip - GlobalSQA").Frame("Frame").WebEdit("address").GetROProperty("title")

' Verify the tooltip text
Dim result2
If tooltipText2 = "Your home or work address." Then
    result2 = "Passed"
Else
    result2 = "Failed"
End If

' Write result to Excel
WriteResultToExcel "DTSToolTip_FormBased", "Verify Tooltip of a adress textbox", "Your home or work address.", tooltipText2, result2, "G:\UFT Projects\GlobalSqa_Automation_Project\Results\Reports\TestReport.xlsx"


Browser("Tooltip - GlobalSQA_2").Page("Tooltip - GlobalSQA").Frame("Frame").WebEdit("address").Set "hgfhf" @@ script infofile_;_ZIP::ssf8.xml_;_
Browser("Tooltip - GlobalSQA_2").Page("Tooltip - GlobalSQA").Frame("Frame").WebButton("Show help").Click @@ script infofile_;_ZIP::ssf9.xml_;_
Browser("Tooltip - GlobalSQA_2").Page("Tooltip - GlobalSQA").WebElement("Forms Based").Click
Browser("Tooltip - GlobalSQA_2").Close @@ hightlight id_;_790288_;_script infofile_;_ZIP::ssf11.xml_;_
