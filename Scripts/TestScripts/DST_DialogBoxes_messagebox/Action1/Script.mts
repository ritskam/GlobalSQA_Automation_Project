Browser("Home - GlobalSQA").Page("Home - GlobalSQA").Link("Dialog Boxes").Click @@ script infofile_;_ZIP::ssf9.xml_;_
Browser("Home - GlobalSQA").Page("Dialog Boxes - GlobalSQA").WebElement("Message Box").Click @@ script infofile_;_ZIP::ssf10.xml_;_
Browser("Home - GlobalSQA").Page("Dialog Boxes - GlobalSQA_2").Frame("Frame").WebElement("ui-id-1").Click @@ hightlight id_;_1376346_;_script infofile_;_ZIP::ssf18.xml_;_

'create a variable and store the identified element 
Set messageBoxAppears = Browser("Home - GlobalSQA").Page("Dialog Boxes - GlobalSQA_2").Frame("Frame").WebElement("Your files have downloaded")

'Check if message box appears
If messageBoxAppears.Exist(5)  Then
 'message appears successfully 
 result = "Passed"
else
  result = "Failed"
End If

'Output the result
WriteResulttoExcel "DST_DialogBoxes_messagebox", "Verify if message box appears", "Your files have downloaded successfully into the My Downloads folder. ", messageBoxAppears.GetRoProperty("innertext") , result, "G:\UFT Projects\GlobalSqa_Automation_Project\Results\Reports\TestReport.xlsx"
Browser("Home - GlobalSQA").Page("Dialog Boxes - GlobalSQA_2").Frame("Frame").WebButton("Ok").Click @@ script infofile_;_ZIP::ssf21.xml_;_
Browser("Home - GlobalSQA").CloseAllTabs

