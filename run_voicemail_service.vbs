' Silent launcher for Voicemail Transcriber Service
' No console window is shown. Works from any location (Startup folder, etc.)

Set WshShell = CreateObject("WScript.Shell")

projectDir = "C:\CLAUDE CODE DESKTOP PROJECTS\Voicemail transcriber"
WshShell.CurrentDirectory = projectDir
WshShell.Run "python """ & projectDir & "\voicemail_service.py""", 0, False
