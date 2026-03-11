Option Explicit

Dim shell
Dim fso
Dim rootDir
Dim command

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

rootDir = fso.GetParentFolderName(WScript.ScriptFullName)
command = "pythonw """ & rootDir & "\server.py"""

' Start backend without opening a console window.
shell.Run command, 0, False

' 1212Give server a short moment to bind the port, then open frontend.
WScript.Sleep 1200
shell.Run "http://127.0.0.1:9527", 0, False


