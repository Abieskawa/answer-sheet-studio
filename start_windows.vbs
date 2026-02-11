Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
repo = fso.GetParentFolderName(WScript.ScriptFullName)

q = Chr(34)
launcher = repo & "\launcher_headless.py"
outputsDir = repo & "\outputs"
startLogPath = outputsDir & "\start_windows.log"

Sub EnsureOutputsDir()
    On Error Resume Next
    If Not fso.FolderExists(outputsDir) Then fso.CreateFolder outputsDir
End Sub

Sub LogLine(msg)
    On Error Resume Next
    EnsureOutputsDir
    Dim f
    Set f = fso.OpenTextFile(startLogPath, 8, True)
    f.WriteLine "[" & Now() & "] " & msg
    f.Close
End Sub

Function ReadTextFile(path)
    On Error Resume Next
    Dim f, t
    ReadTextFile = ""
    If Not fso.FileExists(path) Then Exit Function
    Set f = fso.OpenTextFile(path, 1, False)
    t = f.ReadAll
    f.Close
    ReadTextFile = Trim(CStr(t))
End Function

Function IsPortOpen(port)
    On Error Resume Next
    Dim result
    ' Use PowerShell for reliable port check
    result = WshShell.Run("powershell -NoProfile -ExecutionPolicy Bypass -Command " & q & "$c = New-Object System.Net.Sockets.TcpClient; try { $c.Connect('127.0.0.1', " & port & "); exit 0 } catch { exit 1 } finally { if ($c) { $c.Dispose() } }" & q, 0, True)
    IsPortOpen = (result = 0)
End Function

Function CanRun(cmd)
    On Error Resume Next
    Dim rc
    rc = WshShell.Run(cmd, 0, True)
    CanRun = (Err.Number = 0 And rc = 0)
    Err.Clear
End Function

Sub RunAsync(cmd)
    On Error Resume Next
    LogLine "Running: " & cmd
    WshShell.Run cmd, 0, False
End Sub

Function WrapLauncherCmd(cmd)
    ' Force OPEN_BROWSER=0 via process environment variable instead of cmd /c
    WshShell.Environment("Process")("ANSWER_SHEET_OPEN_BROWSER") = "0"
    WrapLauncherCmd = cmd
End Function

Sub WaitAndOpenUrl(url, port)
    On Error Resume Next
    Dim i
    LogLine "Waiting for port " & port & " (URL: " & url & ")"
    For i = 1 To 150 ' Wait up to 15 seconds for the port to open
        If IsPortOpen(port) Then
            LogLine "Port " & port & " is open. Launching browser..."
            WScript.Sleep 500
            WshShell.Run url, 1, False
            Exit Sub
        End If
        WScript.Sleep 100
    Next
    LogLine "Timeout waiting for port " & port & ". Opening anyway."
    WshShell.Run url, 1, False
End Sub

Sub LaunchThroughLauncher(pyCmd)
    On Error Resume Next
    Dim urlPath, i, u, pParts, pPort
    urlPath = outputsDir & "\progress_url.txt"
    
    ' Clean up old URL file
    If fso.FileExists(urlPath) Then fso.DeleteFile(urlPath)
    
    ' Start launcher
    RunAsync WrapLauncherCmd(pyCmd & " " & q & launcher & q)
    
    ' Wait for the progress server to generate its URL
    LogLine "Waiting for progress_url.txt..."
    For i = 1 To 200 ' 20 seconds
        u = ReadTextFile(urlPath)
        If u <> "" Then
            LogLine "Progress URL found: " & u
            pParts = Split(u, ":")
            If UBound(pParts) >= 2 Then
                pPort = Split(pParts(2), "/")(0)
                WaitAndOpenUrl u, pPort
            Else
                WshShell.Run u, 1, False
            End If
            Exit Sub
        End If
        WScript.Sleep 100
    Next
    LogLine "Error: Launcher progress page did not appear."
End Sub

' --- Main Logic ---
LogLine "start_windows.vbs started."

' 1. If app is already up, just open it
If IsPortOpen(8000) Then
    LogLine "App already running on port 8000. Opening browser."
    WshShell.Run "http://127.0.0.1:8000", 1, False
    WScript.Quit
End If

' 2. Find a suitable Python and launch
probe = " -c " & q & "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)" & q

' Priority 1: Generic pythonw (often the most direct path)
If CanRun("pythonw" & probe) Then
    LaunchThroughLauncher("pythonw")
    WScript.Quit
End If

' Priority 2: Generic pyw
If CanRun("pyw" & probe) Then
    LaunchThroughLauncher("pyw")
    WScript.Quit
End If

' Priority 3: Python Launcher (py -3w)
If CanRun("py -3w" & probe) Then
    LaunchThroughLauncher("py -3w")
    WScript.Quit
End If

' Priority 4: Specific versions
pyVers = Array("3.11", "3.10", "3.12")
For Each v In pyVers
    cmd = "pyw -" & v
    If CanRun(cmd & probe) Then
        LaunchThroughLauncher(cmd)
        WScript.Quit
    End If
Next

' Fallback: Tell the user to install Python
LogLine "No suitable Python 3.10+ found."
WshShell.Popup "Python 3.10+ was not found. Please install Python 3.11 or later from python.org.", 0, "Answer Sheet Studio", 48
WshShell.Run "https://www.python.org/downloads/windows/", 1, False
WScript.Quit 1
