' Voice Assistant Admin Launcher - Completely Hidden
' This VBScript runs the Electron app with admin privileges and no visible console

Set objShell = CreateObject("Shell.Application")
Set fso = CreateObject("Scripting.FileSystemObject")

' Get the directory where this script is located
scriptPath = WScript.ScriptFullName
scriptFolder = fso.GetFile(scriptPath).ParentFolder.Path
parentFolder = fso.GetFolder(scriptFolder).ParentFolder.Path

' Create a temporary batch file that runs npm start
tempBatFile = fso.GetSpecialFolder(2) & "\npm_start_" & Int(Rnd * 1000000) & ".bat"

' Write batch file content
Set batFile = fso.CreateTextFile(tempBatFile, True)
batFile.WriteLine("@echo off")
batFile.WriteLine("cd /d """ & parentFolder & """")
batFile.WriteLine("npm start")
batFile.Close()

' Launch the batch file with admin privileges (hidden window)
objShell.ShellExecute tempBatFile, "", "", "runas", 0

' Give it a moment to start, then delete the batch file
WScript.Sleep(500)
On Error Resume Next
fso.DeleteFile(tempBatFile)
On Error Goto 0

' Exit silently
WScript.Quit(0)
