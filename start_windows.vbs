Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
repo = fso.GetParentFolderName(WScript.ScriptFullName)

cmd = "pythonw " & Chr(34) & repo & "\launcher_gui.py" & Chr(34)
WshShell.Run cmd, 0, False
