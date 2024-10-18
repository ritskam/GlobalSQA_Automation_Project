' Include the SRFExecution.vbs file
ExecuteFile "G:\UFT Projects\GlobalSqa_Automation_Project\Functions\ReusableFunctions\SRFExecution.vbs"

' Specify the path to the SRF executable
Dim srfExecutable
srfExecutable = "G:\UFT Projects\GlobalSqa_Automation_Project\Functions\ReusableFunctions\Executable.exe"

' List of test scripts
Dim testScripts
testScripts = Array( _
    "C:\Path\To\TestScript1.tsp", _
    "C:\Path\To\TestScript2.tsp", _
    "C:\Path\To\TestScript3.tsp", _
    "C:\Path\To\TestScript4.tsp", _
    "C:\Path\To\TestScript5.tsp", _
    "C:\Path\To\TestScript6.tsp", _
    "C:\Path\To\TestScript7.tsp", _
    "C:\Path\To\TestScript8.tsp", _
    "C:\Path\To\TestScript9.tsp", _
    "C:\Path\To\TestScript10.tsp" _
)

' Run each test script in parallel
For Each srfTestPath In testScripts
    ' You can modify this to actually run in parallel if needed
    RunSRFTest srfExecutable, srfTestPath
Next
