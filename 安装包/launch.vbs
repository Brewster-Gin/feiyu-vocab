' Launch reminder.ps1 without any console window.
'   wscript launch.vbs            -> normal run (respects "already done today")
'   wscript launch.vbs preview    -> force show the popup
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
base = fso.GetParentFolderName(WScript.ScriptFullName)
target = base & "\reminder.ps1"

extra = ""
If WScript.Arguments.Length > 0 Then
    If LCase(WScript.Arguments(0)) = "preview" Then extra = " -Preview"
End If

sh.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & target & """" & extra, 0, False
