Set WShell = CreateObject("WScript.Shell")
Dim scriptDir
scriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

WShell.Run "python """ & scriptDir & "\app.py""", 0, False

WScript.Sleep 1500

Do
    WScript.Sleep 2000
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "http://localhost:5001/health", False
    http.Send
    If http.Status = 200 Then
        WShell.Run "http://localhost:5001", 1, False
        Exit Do
    End If
    On Error GoTo 0
Loop
