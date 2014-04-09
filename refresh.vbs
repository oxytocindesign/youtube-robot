On Error Resume Next

Set objExplorer = CreateObject("InternetExplorer.Application")

objExplorer.Navigate "www.abcmouse.com"   
objExplorer.Visible = 1

Wscript.Sleep 3000

Set objDoc = objExplorer.Document

Do While True
    Wscript.Sleep 3000
    objDoc.Location.Reload(True)
    If Err <> 0 Then
        Wscript.Quit
    End If
Loop