Function SelfClosingMsgBox(message, timeout, title)
    Set wshShell = CreateObject(WScript.Shell)
    wshShell.Popup message, timeout, title
    Set wshShell = Nothing
End Function