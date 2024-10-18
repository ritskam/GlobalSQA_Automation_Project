' Load the Excel Data into the Global Sheet
DataTable.ImportSheet "G:\UFT Projects\GlobalSqa_Automation_Project\TestData\InputDataForRigistrationForm.xlsx", 1, dtGlobalSheet

' Get the number of rows in the data sheet
numRows = DataTable.GetSheet(dtGlobalSheet).GetRowCount

Browser("AngularJS User Registration").Navigate "https://globalsqa.com/angularJs-protractor/registration-login-example/#/login" @@ hightlight id_;_10487342_;_script infofile_;_ZIP::ssf9.xml_;_


' Loop through the rows in the Excel sheet
For i = 1 To numRows

    ' Set the current row
    DataTable.SetCurrentRow(i)
    
    
        ' Fetch browser name from the data sheet
    browserName = DataTable("BrowserName", dtGlobalSheet)
    
    ' Call the cross-browser function to initialize the browser
    Set currentBrowser = InitializeBrowser(browserName)


Wait 2 ' Wait for 2 seconds before clicking


Browser("AngularJS User Registration").Page("AngularJS User Registration").WebEdit("username").Set DataTable("Username", dtGlobalSheet) @@ script infofile_;_ZIP::ssf10.xml_;_
Browser("AngularJS User Registration").Page("AngularJS User Registration").WebEdit("password").Set DataTable("Password", dtGlobalSheet) @@ script infofile_;_ZIP::ssf12.xml_;_
Browser("AngularJS User Registration").Page("AngularJS User Registration").WebButton("Login").Click @@ script infofile_;_ZIP::ssf13.xml_;_

confirmationMessage= Browser("AngularJS User Registration").Page("AngularJS User Registration").WebElement("innertext:=You're logged in!!").getROProperty("innertext")
    ' Step 5: Verification
    If confirmationMessage ="You're logged in!!"  Then 
        result = "Passed" & i
    Else
     result = "Failed" & i
    End If    
    
        ' Write result to Excel
 WriteResultToExcel "DTS_login test",  "Verify if user successfully logs in" & i , "You're logged in!!", confirmationMessage, result, "G:\UFT Projects\GlobalSqa_Automation_Project\Results\Reports\TestReportSheet.xlsx"
    
    Browser("AngularJS User Registration").Page("AngularJS User Registration").WebButton("Logout").Click
    
  
Next

Browser("AngularJS User Registration").Close @@ hightlight id_;_855698_;_script infofile_;_ZIP::ssf23.xml_;_
