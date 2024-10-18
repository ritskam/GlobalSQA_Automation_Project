Browser("Home - GlobalSQA_2").Page("Home - GlobalSQA").Link("Registration Form").Click @@ script infofile_;_ZIP::ssf10.xml_;_
Browser("Home - GlobalSQA_2").Page("AngularJS User Registration").WebButton("Register").Click @@ script infofile_;_ZIP::ssf11.xml_;_
Browser("Home - GlobalSQA_2").Page("AngularJS User Registration").WebEdit("firstName").Set "rami" @@ script infofile_;_ZIP::ssf12.xml_;_
Browser("Home - GlobalSQA_2").Page("AngularJS User Registration").WebEdit("lastName").Set "kaki" @@ script infofile_;_ZIP::ssf13.xml_;_
Browser("Home - GlobalSQA_2").Page("AngularJS User Registration").WebEdit("username").Set "ramikaki" @@ script infofile_;_ZIP::ssf14.xml_;_
Browser("Home - GlobalSQA_2").Page("AngularJS User Registration").WebEdit("password").SetSecure "66ffccfa12ac91732431acda74a2eed7249e0c9d8a35" @@ script infofile_;_ZIP::ssf15.xml_;_
Browser("Home - GlobalSQA_2").Page("AngularJS User Registration").WebButton("Register_2").Click @@ script infofile_;_ZIP::ssf16.xml_;_
Browser("Home - GlobalSQA_2").Page("AngularJS User Registration").WebButton("Register").Click @@ script infofile_;_ZIP::ssf17.xml_;_
Browser("Home - GlobalSQA_2").Page("AngularJS User Registration").WebEdit("firstName").Set "sani" @@ script infofile_;_ZIP::ssf18.xml_;_
Browser("Home - GlobalSQA_2").Page("AngularJS User Registration").WebEdit("lastName").Set "kaki" @@ script infofile_;_ZIP::ssf19.xml_;_
Browser("Home - GlobalSQA_2").Page("AngularJS User Registration").WebEdit("username").Set "sanukaki" @@ script infofile_;_ZIP::ssf20.xml_;_
Browser("Home - GlobalSQA_2").Page("AngularJS User Registration").WebEdit("password").SetSecure "66ffcd0fce302e70203e8345dda56913e43a70997e6c" @@ script infofile_;_ZIP::ssf22.xml_;_
Browser("Home - GlobalSQA_2").Page("AngularJS User Registration").WebButton("Register_2").Click @@ script infofile_;_ZIP::ssf23.xml_;_
Browser("Home - GlobalSQA_2").Page("AngularJS User Registration").Sync
Browser("Home - GlobalSQA_2").Close




''Browser("Home - GlobalSQA").Page("Home - GlobalSQA").Link("Registration Form").Click @@ script infofile_;_ZIP::ssf1.xml_;_
''Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebButton("Register").Click @@ script infofile_;_ZIP::ssf2.xml_;_
''Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("firstName").Set "Rutika" @@ script infofile_;_ZIP::ssf3.xml_;_
''Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("lastName").Set "Kamble" @@ script infofile_;_ZIP::ssf4.xml_;_
''Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("username").Set "rutika1" @@ script infofile_;_ZIP::ssf5.xml_;_
''Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("password").SetSecure "66ffb59bd9579f2a88645ed56aa7034ce00b25df5c6c" @@ script infofile_;_ZIP::ssf6.xml_;_
''Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebButton("Register_2").Click @@ script infofile_;_ZIP::ssf9.xml_;_
''Browser("Home - GlobalSQA").Close
'
'' Data-driven Registration Test Script
 @@ hightlight id_;_1771634_;_script infofile_;_ZIP::ssf24.xml_;_
'
'' Load the Excel Data
'DataTable.ImportSheet "G:\UFT Projects\GlobalSqa_Automation_Project\TestData\InputDataForRigistrationForm.xlsx", 1, dtGlobalSheet
'
'' Loop through all rows in the data table
'For i = 1 To DataTable.GetRowCount()
'    
'    ' Set the current row
'    DataTable.SetCurrentRow(i)
'    
'    ' Step 1: Navigate to the Registration Page
'    Browser("Home - GlobalSQA").Page("Home - GlobalSQA").Link("Registration Form").Click
'    Browser("AngularJS User Registration").Page("AngularJS User Registration").WebButton("Register").Click
'    
'    ' Step 2: Fill out the Registration Form
'    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("firstName").Set DataTable("FirstName", dtGlobalSheet)
'    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("lastName").Set DataTable("LastName", dtGlobalSheet)
'    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("username").Set DataTable("Username", dtGlobalSheet)
'    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("password").SetSecure DataTable("Password", dtGlobalSheet)
'    
'    ' Step 3: Click Register
'    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebButton("Register_2").Click
'    
'    ' Step 4: Verify if Registration Passed or Failed
'    If Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebElement("Registration successful").Exist(5) Then
'        Reporter.ReportEvent micPass, "Registration Test", "User registered successfully for row: " & i
'    Else
'        Reporter.ReportEvent micFail, "Registration Test", "Registration failed for row: " & i
'    End If
'    
'    ' Step 5: Close the Browser
'    Browser("Home - GlobalSQA").Close
'Next
'
'
'
''
''
''
''
''
''
''
'
'
'
'
'
'
'
'
'
'
'
'
'
'
