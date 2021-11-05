Option Explicit

Dim ws, fso, objShell
Set ws = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

Dim lnk, resPath, nbPath, command

Set lnk = ws.Createshortcut(WScript.Arguments(0))
nbPath = """" + fso.GetParentFolderName(fso.GetParentFolderName(lnk.TargetPath)) + "\"""

resPath = """" + fso.GetParentFolderName(Wscript.ScriptFullName) + "\res\"""

command = "/c"
command = command + " xcopy " + resPath + " " + nbPath + " /S /Y"
command = command + " && timeout 5"

objShell.ShellExecute "cmd.exe", command, , "runas", 1
