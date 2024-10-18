' Declare variables
Dim objExcel, objSheet
Dim rowCount, username, password, expectedResult

' Create an instance of Excel application
Set objExcel = CreateObject("Excel.Application")

' Open the Excel file
Set objWorkbook = objExcel.Workbooks.Open("C:\UFT\GlobalSQA_Automation\Data\LoginData.xlsx")

' Set the worksheet (assuming data is on the first sheet)
Set objSheet = objWorkbook.Sheets(1)

' Get the number of rows with data (assuming no empty rows between data)
rowCount = objSheet.UsedRange.Rows.Count

' Loop through the rows (starting from row 2 to skip the header)
For i = 2 To rowCount
    ' Read the data from the Excel file
    username = objSheet.Cells(i, 1).Value
    password = objSheet.Cells(i, 2).Value
    expectedResult = objSheet.Cells(i, 3).Value
    
    ' Call the login function with Excel data
    LoginToGlobalSQA username, password
    
    ' Add validation step here to check if the result matches expectedResult
    If CheckLoginResult() = expectedResult Then
        Reporter.ReportEvent micPass, "Login Test", "Passed for: " & username
    Else
        Reporter.ReportEvent micFail, "Login Test", "Failed for: " & username
    End If
Next

' Close the workbook and quit Excel
objWorkbook.Close False
objExcel.Quit

' Clean up the objects
Set objSheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
