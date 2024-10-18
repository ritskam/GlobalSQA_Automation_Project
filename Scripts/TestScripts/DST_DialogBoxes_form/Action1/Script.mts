Browser("Home - GlobalSQA").Page("Home - GlobalSQA").Link("Dialog Boxes").Click
Browser("Home - GlobalSQA").Page("Dialog Boxes - GlobalSQA").Frame("Frame").WebButton("Create new user").Click
Browser("Home - GlobalSQA").Page("Dialog Boxes - GlobalSQA").Frame("Frame").WebElement("Create new user").Click
Browser("Home - GlobalSQA").Page("Dialog Boxes - GlobalSQA").Frame("Frame").WebEdit("name").Set "ritz kamle"
Browser("Home - GlobalSQA").Page("Dialog Boxes - GlobalSQA").Frame("Frame").WebElement("Existing Users: Name Email").Click
Browser("Home - GlobalSQA").Page("Dialog Boxes - GlobalSQA").Frame("Frame").WebEdit("email").Set "ritz@gmail.com"
Browser("Home - GlobalSQA").Page("Dialog Boxes - GlobalSQA").Frame("Frame").WebElement("Existing Users: Name Email").Click
Browser("Home - GlobalSQA").Page("Dialog Boxes - GlobalSQA").Frame("Frame").WebEdit("password").SetSecure "66fd7edfe1125aa8d52bda9d7f336b87bbda"
Browser("Home - GlobalSQA").Page("Dialog Boxes - GlobalSQA").Frame("Frame").WebButton("Create an account").Click

' Verify if the new user is added to the table
Set newUserEmail = Browser("Home - GlobalSQA").Page("Dialog Boxes - GlobalSQA").Frame("Frame").WebElement("innertext:=ritz@gmail.com")
Set newUserName = Browser("Home - GlobalSQA").Page("Dialog Boxes - GlobalSQA").Frame("Frame").WebElement("innertext:=ritz kamle")

' Check if new user email and name are present
If newUserEmail.Exist(5) And newUserName.Exist(5) Then
    ' User is successfully added
    result = "Passed"
Else
    ' User is not found
    result = "Failed"
End If

' Output the result
WriteResulttoExcel "new_user_check", "Verify new user addition", "Expected to find ritz@gmail.com and ritz kamle", newUserEmail.GetRoProperty("innertext") & " and " & newUserName.GetRoProperty("innertext"), result, "G:\UFT Projects\GlobalSqa_Automation_Project\Results\Reports\TestReport.xlsx"

Browser("Home - GlobalSQA").CloseAllTabs

