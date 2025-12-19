Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
repo = fso.GetParentFolderName(WScript.ScriptFullName)

q = Chr(34)
launcher = repo & "\launcher_gui.py"

' Prefer supported Python versions (3.10–3.13) via the Windows py launcher if available.
cmd = "cmd /c " & q & "(" & _
    "pyw -3.13 " & q & launcher & q & " || " & _
    "pyw -3.12 " & q & launcher & q & " || " & _
    "pyw -3.11 " & q & launcher & q & " || " & _
    "pyw -3.10 " & q & launcher & q & " || " & _
    "pythonw " & q & launcher & q & _
    ")" & q
WshShell.Run cmd, 0, False
