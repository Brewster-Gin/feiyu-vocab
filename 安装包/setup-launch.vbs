' Feiyu Vocab installer launcher: starts the GUI with no console window.
Dim fso, sh, dir, cmd, rc
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ _
    & dir & "\setup-gui.ps1"""
rc = sh.Run(cmd, 0, True)
If rc <> 0 Then
    MsgBox "The installer could not start (exit code " & rc & ")." & vbCrLf & vbCrLf & _
           "Log: " & sh.ExpandEnvironmentStrings("%TEMP%") & "\feiyu-install.log", _
           16, "Feiyu Vocab Setup"
End If
