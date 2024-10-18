' SRFExecution.vbs
Function RunSRFTest(srfExecutable, srfTestPath)
    Dim objShell
    Set objShell = CreateObject("WScript.Shell")
    ' Execute the SRF test
    objShell.Run Chr(34) & srfExecutable & Chr(34) & " " & Chr(34) & srfTestPath & Chr(34), 1, True
    Set objShell = Nothing
End Function
