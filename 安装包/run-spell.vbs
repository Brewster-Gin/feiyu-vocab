' Launch the spelling test without any console window.
'   wscript run-spell.vbs <day> [test]
' Returns the exit code of spell.ps1 (0 = passed, 1 = round void, 2 = cancelled).
Dim fso, sh, base, script, day, extra, cmd, code
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
base = fso.GetParentFolderName(WScript.ScriptFullName)
script = base & "\spell.ps1"

day = "1"
If WScript.Arguments.Length > 0 Then day = WScript.Arguments(0)

extra = ""
If WScript.Arguments.Length > 1 Then
    If LCase(WScript.Arguments(1)) = "test" Then extra = " -Test"
End If

cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & script & """ -Day " & day & extra
code = sh.Run(cmd, 0, True)
WScript.Quit code
