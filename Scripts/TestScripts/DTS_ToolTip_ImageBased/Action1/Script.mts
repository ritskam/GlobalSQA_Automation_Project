Browser("Tooltip - GlobalSQA").Page("Tooltip - GlobalSQA").Sync
Browser("Tooltip - GlobalSQA").Navigate "https://www.globalsqa.com/"
Browser("Tooltip - GlobalSQA").Page("Home - GlobalSQA").Link("Tooltip").Click


' Usage of the function in your script
SelfClosingMsgBox "Vienna, Austria", 2, "Tooltip Text"


' Store the tooltip text in a variable
tooltipText = Browser("Tooltip - GlobalSQA").Page("Tooltip - GlobalSQA").Frame("Frame").Link("Vienna, Austria").GetRoProperty("text")

' Verify the tooltip text
Dim result
If tooltipText = "Vienna, Austria" Then
    result = "Passed"
Else
    result = "Failed"
End If

' Write result to Excel
WriteResultToExcel "tooltip_test", "Verify Tooltip Text", "Vienna, Austria", tooltipText, result, "G:\UFT Projects\GlobalSqa_Automation_Project\Results\Reports\TestReport.xlsx"

 @@ script infofile_;_ZIP::ssf3.xml_;_
' Usage of the function in your script
SelfClosingMsgBox "London, England", 2, "Tooltip Text"

' Store the tooltip text in a variable
tooltipText1 = Browser("Tooltip - GlobalSQA").Page("Tooltip - GlobalSQA").Frame("Frame").Link("London, England").GetRoProperty("text")

' Verify the tooltip text
Dim result1
If tooltipText1 = "London, England" Then
    result1 = "Passed"
Else
    result1 = "Failed"
End If

' Write result to Excel
WriteResultToExcel "tooltip_test_text2", "Verify Tooltip Text of text", "London, England", tooltipText1, result1, "G:\UFT Projects\GlobalSqa_Automation_Project\Results\Reports\TestReport.xlsx"
 @@ script infofile_;_ZIP::ssf7.xml_;_


' Usage of the function in your script
SelfClosingMsgBox "St", 2, "Tooltip Text"


' Store the tooltip text in a variable
tooltipText2 = Browser("Tooltip - GlobalSQA").Page("Tooltip - GlobalSQA").Frame("Frame").Image("St").GetRoProperty("alt")

' Verify the tooltip text
Dim result2
If tooltipText2 = "St. Stephen's Cathedral" Then
    result2 = "Passed"
Else
    result2 = "Failed"
End If

' Write result to Excel
WriteResultToExcel "tooltip_test1", "Verify Tooltip test of an image", "St. Stephen's Cathedral", tooltipText2, result2, "G:\UFT Projects\GlobalSqa_Automation_Project\Results\Reports\TestReport.xlsx"
 @@ script infofile_;_ZIP::ssf9.xml_;_

' Usage of the function in your script
SelfClosingMsgBox "Tower Bridge", 2, "Tooltip Text3"


' Store the tooltip text in a variable
tooltipText3 = Browser("Tooltip - GlobalSQA").Page("Tooltip - GlobalSQA").Frame("Frame").Image("Tower Bridge").GetRoProperty("alt")

' Verify the tooltip text
Dim result3
If tooltipText3 = "Tower Bridge" Then
    result3 = "Passed"
Else
    result3 = "Failed"
End If

' Write result to Excel
WriteResultToExcel "tooltip_test_image2", "Verify Tooltip Text of an image", "Tower Bridge", tooltipText3, result3, "G:\UFT Projects\GlobalSqa_Automation_Project\Results\Reports\TestReport.xlsx"

Browser("Tooltip - GlobalSQA").CloseAllTabs @@ hightlight id_;_4588306_;_script infofile_;_ZIP::ssf11.xml_;_


