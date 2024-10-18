'Browser("Home - GlobalSQA").Page("Home - GlobalSQA").Link("Registration Form").Click @@ script infofile_;_ZIP::ssf1.xml_;_
'Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebButton("Register").Click @@ script infofile_;_ZIP::ssf2.xml_;_
'Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("firstName").Set "pepa" @@ script infofile_;_ZIP::ssf3.xml_;_
'Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("lastName").Set "pg" @@ script infofile_;_ZIP::ssf4.xml_;_
'Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("username").Set "pepapig" @@ script infofile_;_ZIP::ssf5.xml_;_
'Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("password").SetSecure "66ffcf2e27f0cc80c4b0f4592e10b4fe843a7f470240798c" @@ script infofile_;_ZIP::ssf6.xml_;_
'Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebButton("Register_2").Click @@ script infofile_;_ZIP::ssf7.xml_;_
'Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebButton("Register").Click @@ script infofile_;_ZIP::ssf8.xml_;_
'Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("firstName").Set "sin" @@ script infofile_;_ZIP::ssf9.xml_;_
'Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("lastName").Set "can" @@ script infofile_;_ZIP::ssf10.xml_;_
'Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("username").Set "sincan" @@ script infofile_;_ZIP::ssf11.xml_;_
'Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("password").SetSecure "66ffcf419dd0e4ff7d634d3f0cc454d6aaf6f7df8149" @@ script infofile_;_ZIP::ssf12.xml_;_
'Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebButton("Register_2").Click @@ script infofile_;_ZIP::ssf13.xml_;_
'Browser("Home - GlobalSQA").Page("AngularJS User Registration").Sync
'Browser("Home - GlobalSQA").Close @@ hightlight id_;_1771634_;_script infofile_;_ZIP::ssf14.xml_;_

' Load the Excel Data into the Data Table
DataTable.ImportSheet "G:\UFT Projects\GlobalSqa_Automation_Project\TestData\InputDataForRigistrationForm.xlsx", 1, dtGlobalSheet

' Loop through all rows in the data table
For i = 1 To DataTable.GetRowCount()
    
    ' Set the current row
    DataTable.SetCurrentRow(i)
    
    ' Step 1: Open a new browser for each iteration and navigate to the Registration Page
    SystemUtil.Run "iexplore.exe", "https://globalsqa.com/angularJs-protractor/registration-login-example/#/login"
    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebButton("Register").Click
    
    ' Step 2: Fill out the Registration Form with data from Excel
    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("firstName").Set DataTable("FirstName", dtGlobalSheet)
    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("lastName").Set DataTable("LastName", dtGlobalSheet)
    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("username").Set DataTable("Username", dtGlobalSheet)
    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("password").SetSecure DataTable("Password", dtGlobalSheet)

    ' Step 3: Click the Register button
    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebButton("Register_2").Click
    
    ' Step 4: Verify if Registration Passed or Failed
    If Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebElement("Registration successful").Exist(5) Then
        Reporter.ReportEvent micPass, "Registration Test", "User registered successfully for row: " & i
    Else
        Reporter.ReportEvent micFail, "Registration Test", "Registration failed for row: " & i
    End If
    
    ' Step 5: Close the Browser
    Browser("Home - GlobalSQA").Close

Next

