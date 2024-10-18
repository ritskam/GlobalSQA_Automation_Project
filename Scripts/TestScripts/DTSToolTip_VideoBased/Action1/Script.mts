
Browser("Home - GlobalSQA").Page("Home - GlobalSQA").Link("Tooltip").Click @@ script infofile_;_ZIP::ssf7.xml_;_
Browser("Home - GlobalSQA").Page("Tooltip - GlobalSQA").WebElement("Video Based").Click @@ script infofile_;_ZIP::ssf8.xml_;_


' Usage of the function in your script
SelfClosingMsgBox "Like", 2, "Tooltip Text"

' Store the tooltip text in a variable
tooltipText = Browser("Tooltip - GlobalSQA").Page("Tooltip - GlobalSQA").Frame("Frame").WebButton("Like").GetROProperty("title")

' Verify the tooltip text
Dim result
If tooltipText = "I like this" Then
    result = "Passed"
Else
    result = "Failed"
End If

' Write result to Excel
WriteResultToExcel "DTSToolTip_VideoBased", "Verify Tooltip of like button", "I like this", tooltipText, result, "G:\UFT Projects\GlobalSqa_Automation_Project\Results\Reports\TestReport.xlsx"


Browser("Tooltip - GlobalSQA").Page("Tooltip - GlobalSQA").Frame("Frame").WebButton("Like").Click @@ script infofile_;_ZIP::ssf9.xml_;_



' Usage of the function in your script
SelfClosingMsgBox "Add to", 2, "Tooltip Text"

' Store the tooltip text in a variable
tooltipText1 = Browser("Tooltip - GlobalSQA").Page("Tooltip - GlobalSQA").Frame("Frame").WebButton("Add to").GetROProperty("title")

' Verify the tooltip text
Dim result1
If tooltipText1 = "Add to Watch Later" Then
    result1 = "Passed"
Else
    result1 = "Failed"
End If

' Write result to Excel
WriteResultToExcel "DTSToolTip_VideoBased", "Verify Tooltip of Add to button", "Add to Watch Later", tooltipText1, result1, "G:\UFT Projects\GlobalSqa_Automation_Project\Results\Reports\TestReport.xlsx"


Browser("Tooltip - GlobalSQA").Page("Tooltip - GlobalSQA").Frame("Frame").WebButton("Add to").Click @@ script infofile_;_ZIP::ssf10.xml_;_



' Usage of the function in your script
SelfClosingMsgBox "Share", 2, "Tooltip Text"

' Store the tooltip text in a variable
tooltipText2 = Browser("Tooltip - GlobalSQA").Page("Tooltip - GlobalSQA").Frame("Frame").WebButton("Share").GetROProperty("title")

' Verify the tooltip text
Dim result2
If tooltipText2 = "Share this video" Then
    result2 = "Passed"
Else
    result2 = "Failed"
End If

' Write result to Excel
WriteResultToExcel "DTSToolTip_VideoBased", "Verify Tooltip of Share button", "Share this video", tooltipText2, result2, "G:\UFT Projects\GlobalSqa_Automation_Project\Results\Reports\TestReport.xlsx"


'Browser("Home - GlobalSQA").Page("Tooltip - GlobalSQA").Frame("Frame").WebButton("Share").Click @@ script infofile_;_ZIP::ssf11.xml_;_
'Browser("Home - GlobalSQA").Page("Tooltip - GlobalSQA").Frame("Frame").WebButton("I dislike this").Click @@ script infofile_;_ZIP::ssf12.xml_;_
 @@ script infofile_;_ZIP::ssf13.xml_;_
Browser("Home - GlobalSQA").Page("Tooltip - GlobalSQA").Sync
Browser("Home - GlobalSQA").Close @@ hightlight id_;_1119106_;_script infofile_;_ZIP::ssf14.xml_;_
