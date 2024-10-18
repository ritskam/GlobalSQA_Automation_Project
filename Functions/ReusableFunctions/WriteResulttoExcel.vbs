Function WriteResulttoExcel(ID, TestResult, SheetPath)
  'Creating the Excel Object
  set objExcel = createobject(excel.application)
  'Creating the Workbooks object
  set objWB = objExcel.workbooks.open (SheetPath)
  'Creating the sheet object
  set objsheet = objwb.worksheets(1)
  ' Write test results to excel sheet
  rws=objsheet.UsedRange.Rows.count
  objsheet.cells(1,rws+1).Value= ID
  objsheet.cells(2,rws+1).Value= TestResult
  'Saving the workbook after changes
  objWb.save
  'closing the workbook
  objWB.close
 'Quit the Excel and destroying the Excel object
  objExcel.Quit
  set objExcel=nothing
End Function