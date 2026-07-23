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

Function U(h)
    Dim i, s, val
    s = ""
    For i = 1 To Len(h) Step 4
        val = CLng("&H" & Mid(h, i, 4))
        s = s & ChrW(val)
    Next
    U = s
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
probe = " -c " & q & "import sys; raise SystemExit(0 if (3, 10) <= sys.version_info[:2] <= (3, 11) else 1)" & q

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
pyVers = Array("3.11", "3.10")
For Each v In pyVers
    cmd = "pyw -" & v
    If CanRun(cmd & probe) Then
        LaunchThroughLauncher(cmd)
        WScript.Quit
    End If
Next

' Fallback: Tell the user to install Python
LogLine "No suitable Python 3.10 or 3.11 found."
Dim recommendedPythonVersion
recommendedPythonVersion = "3.11.8"
Dim pkgUrl, pkgPath
pkgUrl = "https://www.python.org/ftp/python/" & recommendedPythonVersion & "/python-" & recommendedPythonVersion & "-amd64.exe"
pkgPath = WshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Downloads\answer_sheet_studio_python_" & recommendedPythonVersion & ".exe"
Dim dlChoice
dlChoice = WshShell.Popup(U("672A50756E2C52307CFB7D715B8988DD00200050007900740068006F006E00200033002E003100300020621600200033002E003100313002") & vbCrLf & vbCrLf & U("662F54267ACB53734E0B8F094E26958B555F00200050007900740068006F006E00200033002E00310031002E003800205B8988DD7A0B5F0FFF1F"), 0, "Answer Sheet Studio", 52)

If dlChoice = 6 Then
    Dim dlResult
    LogLine "Downloading Python installer to: " & pkgPath
    dlResult = WshShell.Run("powershell -NoProfile -ExecutionPolicy Bypass -Command " & q & "try { (New-Object System.Net.WebClient).DownloadFile('" & pkgUrl & "', '" & pkgPath & "'); exit 0 } catch { exit 1 }" & q, 0, True)
    If dlResult = 0 And fso.FileExists(pkgPath) Then
        Dim fileSize
        fileSize = fso.GetFile(pkgPath).Size
        If fileSize > 20000000 Then
            LogLine "Download succeeded (" & fileSize & " bytes). Opening installer."
            WshShell.Run q & pkgPath & q, 1, False
            WshShell.Popup U("0050007900740068006F006E00205B8988DD7A0B5F0F5DF24E0B8F094E26958B555F3002") & vbCrLf & vbCrLf & U("8ACB65BC5B8988DD5B8C62105F8CFF0C518D6B2157F7884C00200041006E0073007700650072002000530068006500650074002000530074007500640069006F3002"), 0, "Answer Sheet Studio", 64
        Else
            LogLine "Downloaded file too small (" & fileSize & " bytes). Opening download URL in browser."
            WshShell.Run pkgUrl
            WshShell.Popup U("5DF2958B555F700F89BD566876F463A59032884C00200050007900740068006F006E00205B8988DD6A944E0B8F093002") & vbCrLf & vbCrLf & U("8ACB65BC5B8988DD5B8C62105F8CFF0C518D6B2157F7884C00200041006E0073007700650072002000530068006500650074002000530074007500640069006F3002"), 0, "Answer Sheet Studio", 64
        End If
    Else
        LogLine "Download failed (result=" & dlResult & "). Opening download URL in browser."
        WshShell.Run pkgUrl
        WshShell.Popup U("5DF2958B555F700F89BD566876F463A59032884C00200050007900740068006F006E00205B8988DD6A944E0B8F093002") & vbCrLf & vbCrLf & U("8ACB65BC5B8988DD5B8C62105F8CFF0C518D6B2157F7884C00200041006E0073007700650072002000530068006500650074002000530074007500640069006F3002"), 0, "Answer Sheet Studio", 64
    End If
Else
    Dim openResult
    openResult = WshShell.Popup(U("6B647A0B5F0F970089815B8988DD00200050007900740068006F006E00200033002E00310031002E00380020624D80FD57F7884C3002") & vbCrLf & vbCrLf & U("9EDE64CA300C78BA5B9A300D53EF76F463A55728700F89BD56684E0B8F095B8988DD6A94FF087121970095B18B8082F165875B9865B97DB29801FF093002"), 0, "Answer Sheet Studio", 1)
    If openResult = 1 Then
        LogLine "User chose to open download URL in browser."
        WshShell.Run pkgUrl
    End If
End If
WScript.Quit 1
