' VBScript wrapper to run the batch file without showing a console window
' This creates a more "app-like" experience on Windows

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the directory where this script is located
strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
strBatchFile = strScriptPath & "\launch_app.bat"

' Check if batch file exists
If Not objFSO.FileExists(strBatchFile) Then
    MsgBox "Error: launch_app.bat not found in:" & vbCrLf & strScriptPath, vbCritical, "Topic Generator"
    WScript.Quit
End If

' Run the batch file in a visible window (user can see progress)
objShell.Run """" & strBatchFile & """", 1, False

