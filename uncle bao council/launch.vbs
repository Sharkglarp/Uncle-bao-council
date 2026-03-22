Option Explicit

Dim shell, fso, appDir, scriptPath, command, launchMode
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

appDir = fso.GetParentFolderName(WScript.ScriptFullName)
scriptPath = appDir & "\UncleBaoCouncil.ps1"
launchMode = ""

If WScript.Arguments.Count > 0 Then
    launchMode = LCase(Trim(WScript.Arguments(0)))
End If

command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File """ & scriptPath & """"
If launchMode = "startup" Or launchMode = "/startup" Then
    command = command & " -StartHidden"
End If

shell.Run command, 0, False
