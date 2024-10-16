' Function for Cross-Browser Execution
Function InitializeBrowser(browserName)
    Select Case browserName
        Case "Chrome"
            SystemUtil.Run "chrome.exe", "https://globalsqa.com/angularJs-protractor/registration-login-example/#/login"
            Set InitializeBrowser = Browser("CreationTime:=0")  ' Return the browser object
        Case "Firefox"
            SystemUtil.Run "firefox.exe", "https://globalsqa.com/angularJs-protractor/registration-login-example/#/login"
            Set InitializeBrowser = Browser("CreationTime:=0")  ' Return the browser object
        Case "IE"
            SystemUtil.Run "iexplore.exe", "https://globalsqa.com/angularJs-protractor/registration-login-example/#/login"
            Set InitializeBrowser = Browser("CreationTime:=0")  ' Return the browser object
        Case Else
            Reporter.ReportEvent micFail, "Invalid Browser", "The specified browser is not supported: " & browserName
            Exit Function
    End Select
End Function

