Randomize
Dim port
port = Int(Rnd() * 8974) + 1025
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run """%programfiles%\iis express\iisexpress"" /path:""" & WScript.Arguments.Item(0) & """ /port:" & CStr(port) & " /systray:true", 0, False
WshShell.Run "http://localhost:" & CStr(port), 9, False
Set WshShell = Nothing
