

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


    ' Step 2: Click Register
    Browser("Home - GlobalSQA").Page("AngularJS User Registration").Link("Register").Click
    
    ' Step 3: Fill out the Registration Form with data from the Excel sheet
    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("firstName").Set DataTable("FirstName", dtGlobalSheet)
    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("lastName").Set DataTable("LastName", dtGlobalSheet)
    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("username").Set DataTable("Username", dtGlobalSheet)
    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebEdit("password").Set DataTable("Password", dtGlobalSheet)
    
    ' Step 4: Click Register and Verify
    Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebButton("Register").Click



confirmationMessage= Browser("Home - GlobalSQA").Page("AngularJS User Registration").WebElement("innertext:=Registration successful").Exist(3)
    ' Step 5: Verification
    If confirmationMessage  Then 
        result = "Passed" & i
    Else
     result = "Failed" & i
    End If
    
    ' Write result to Excel
WriteResultToExcel "AngularJS_PracticeSet_RegistrationForm",  "Verify if user successfully registered for user" &i , "Registration successful.", confirmationMessage, result, "G:\UFT Projects\GlobalSqa_Automation_Project\Results\Reports\TestReport.xlsx"
    
Next



' Close Browser
Browser("Home - GlobalSQA").Close

