Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
repo = fso.GetParentFolderName(WScript.ScriptFullName)

q = Chr(34)
launcher = repo & "\launcher_headless.py"

Function CanRun(cmd)
    On Error Resume Next
    Dim rc
    rc = WshShell.Run(cmd, 0, True)
    If Err.Number <> 0 Then
        Err.Clear
        CanRun = False
        Exit Function
    End If
    CanRun = (rc = 0)
End Function

Sub RunAsync(cmd)
    On Error Resume Next
    WshShell.Run cmd, 0, False
End Sub

probe = q & "import sys; raise SystemExit(0)" & q

' Prefer supported Python versions (3.10+) via the Windows py launcher (pyw) if available.
If CanRun("pyw -3.10 -c " & probe) Then
    RunAsync "pyw -3.10 " & q & launcher & q
    WScript.Quit
End If
If CanRun("pyw -3.11 -c " & probe) Then
    RunAsync "pyw -3.11 " & q & launcher & q
    WScript.Quit
End If
If CanRun("pyw -3.12 -c " & probe) Then
    RunAsync "pyw -3.12 " & q & launcher & q
    WScript.Quit
End If
If CanRun("pyw -3.13 -c " & probe) Then
    RunAsync "pyw -3.13 " & q & launcher & q
    WScript.Quit
End If

' Fallback: pythonw (no console window)
If CanRun("pythonw -c " & probe) Then
    RunAsync "pythonw " & q & launcher & q
    WScript.Quit
End If

' Last resort: python (hidden by windowstyle=0)
If CanRun("python -c " & probe) Then
    RunAsync "python " & q & launcher & q
    WScript.Quit
End If

WshShell.Popup "Python not found. Please install Python 3.10+ from python.org.", 0, "Answer Sheet Studio", 48
