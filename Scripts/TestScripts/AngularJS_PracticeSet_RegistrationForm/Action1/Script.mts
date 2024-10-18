' Load the Excel Data into the Global Sheet
DataTable.ImportSheet "G:\UFT Projects\GlobalSqa_Automation_Project\TestData\InputDataForRigistrationForm.xlsx", 1, dtGlobalSheet

' Get the number of rows in the data sheet
numRows = DataTable.GetSheet(dtGlobalSheet).GetRowCount

    ' Step 1: Navigate to the Registration Page
    Browser("Home - GlobalSQA").Navigate "https://globalsqa.com/angularJs-protractor/registration-login-example/#/login"

' Loop through the rows in the Excel sheet
For i = 1 To numRows

    ' Set the current row
    DataTable.SetCurrentRow(i)

Wait 2 ' Wait for 2 seconds before clicking

    ' Step 2: Click Register
    Browser("AngularJS User Registration").Page("AngularJS User Registration").WebButton("Register").Click
  @@ script infofile_;_ZIP::ssf33.xml_;_
    
    ' Step 3: Fill out the Registration Form with data from the Excel sheet
    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("firstName").Set DataTable("FirstName", dtGlobalSheet)
    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("lastName").Set DataTable("LastName", dtGlobalSheet)
    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("username").Set DataTable("Username", dtGlobalSheet)
    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("password").Set DataTable("Password", dtGlobalSheet)
    
    ' Step 4: Click Register and Verify
    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebButton("Register_2").Click
    
    
    

    ' Step 5: Verification
    If Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebElement("innertext:=Registration successful").Exist(5) Then
        Reporter.ReportEvent micPass, "Registration Test", "User registered successfully for row: " & i
    Else
        Reporter.ReportEvent micFail, "Registration Test", "Registration failed for row: " & i
    End If

    ' Sync before moving to the next iteration
    Browser("Home - GlobalSQA").Page("AngularJS User Registration").Sync

Next

' Close Browser
Browser("Home - GlobalSQA").Close


