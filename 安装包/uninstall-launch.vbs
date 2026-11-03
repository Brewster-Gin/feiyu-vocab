' Feiyu Vocab uninstaller launcher: starts the GUI with no console window.
Dim fso, sh, dir, cmd, rc
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ _
    & dir & "\uninstall-gui.ps1"""
rc = sh.Run(cmd, 0, True)
If rc <> 0 Then
    MsgBox "The uninstaller could not start (exit code " & rc & ")." & vbCrLf & vbCrLf & _
           "Log: " & sh.ExpandEnvironmentStrings("%TEMP%") & "\feiyu-uninstall-err.log", _
           16, "Feiyu Vocab Uninstall"
End If
