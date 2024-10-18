' Function for performing the Login test
Function Test_Login()

    ' Step 1: Navigate to Login Page (already navigated by cross-browser function)
    ' You can skip direct navigation here, as it is handled in the RunTestInBrowser function

    ' Step 2: Perform Login Actions (fill the username and password fields)
    Browser("micclass:=Browser").Page("AngularJS User Registration").WebEdit("username").Set DataTable("Username", dtGlobalSheet)
    Browser("micclass:=Browser").Page("AngularJS User Registration").WebEdit("password").Set DataTable("Password", dtGlobalSheet)

    ' Step 3: Click Login Button
    Browser("micclass:=Browser").Page("AngularJS User Registration").WebButton("Login").Click

    ' Step 4: Verify Login Success
    loginMessage = Browser("micclass:=Browser").Page("AngularJS User Registration").WebElement("innertext:=Login successful").Exist(5)
    
    ' Step 5: Report the result of the login action
    If loginMessage Then
        Reporter.ReportEvent micPass, "Login Test", "Login successful"
    Else
        Reporter.ReportEvent micFail, "Login Test", "Login failed"
    End If

End Function
