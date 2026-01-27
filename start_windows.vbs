Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
repo = fso.GetParentFolderName(WScript.ScriptFullName)

q = Chr(34)
launcher = repo & "\launcher_headless.py"
Dim outputsDir, startLogPath
outputsDir = repo & "\outputs"
startLogPath = outputsDir & "\start_windows.log"

Sub EnsureOutputsDir()
    On Error Resume Next
    If Not fso.FolderExists(outputsDir) Then
        fso.CreateFolder outputsDir
    End If
End Sub

Sub LogLine(msg)
    On Error Resume Next
    EnsureOutputsDir
    Dim f
    Set f = fso.OpenTextFile(startLogPath, 8, True)
    f.WriteLine "[" & Now() & "] " & msg
    f.Close
End Sub

LogLine "start_windows.vbs launched, repo=" & repo

recommendedPython = WshShell.Environment("Process")("ANSWER_SHEET_PYTHON_VERSION")
If recommendedPython = "" Then
    recommendedPython = "3.11.8"
End If
LogLine "recommendedPython=" & recommendedPython

Function ReadTextFile(path)
    On Error Resume Next
    Dim f, t
    ReadTextFile = ""
    If Not fso.FileExists(path) Then
        Exit Function
    End If
    Set f = fso.OpenTextFile(path, 1, False)
    t = f.ReadAll
    f.Close
    ReadTextFile = Trim(CStr(t))
End Function

Sub TryOpenProgressPage()
    On Error Resume Next
    Dim outDir, urlPath, i, u
    outDir = repo & "\outputs"
    urlPath = outDir & "\progress_url.txt"
    LogLine "Waiting for progress_url.txt at " & urlPath
    For i = 1 To 50 ' ~5s
        u = ReadTextFile(urlPath)
        If u <> "" Then
            LogLine "Opening progress URL: " & u
            WshShell.Run u, 1, False
            Exit Sub
        End If
        WScript.Sleep 100
    Next
    LogLine "progress_url.txt not found; launcher may not have started."
End Sub

Function CanRun(cmd)
    On Error Resume Next
    Dim rc
    rc = WshShell.Run(cmd, 0, True)
    If Err.Number <> 0 Then
        LogLine "CanRun error (" & Err.Number & "): " & Err.Description & " cmd=" & cmd
        Err.Clear
        CanRun = False
        Exit Function
    End If
    LogLine "CanRun rc=" & rc & " cmd=" & cmd
    CanRun = (rc = 0)
End Function

Sub RunAsync(cmd)
    On Error Resume Next
    LogLine "RunAsync: " & cmd
    WshShell.Run cmd, 0, False
End Sub

Function WrapLauncherCmd(cmd)
    On Error Resume Next
    ' Prevent launcher_headless.py from opening extra progress pages itself.
    WrapLauncherCmd = "cmd.exe /c " & q & "set ANSWER_SHEET_OPEN_BROWSER=0&" & cmd & q
End Function

Function MajorMinorTag(version)
    Dim parts
    parts = Split(version, ".")
    If UBound(parts) < 1 Then
        MajorMinorTag = ""
        Exit Function
    End If
    MajorMinorTag = CStr(parts(0)) & CStr(parts(1))
End Function

Sub TryRunFromKnownInstall(tag)
    On Error Resume Next
    Dim localAppData, programFiles, pythonw, python
    localAppData = WshShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
    programFiles = WshShell.ExpandEnvironmentStrings("%ProgramFiles%")

    pythonw = localAppData & "\Programs\Python\Python" & tag & "\pythonw.exe"
    python = localAppData & "\Programs\Python\Python" & tag & "\python.exe"
    If fso.FileExists(pythonw) Then
        RunAsync WrapLauncherCmd(q & pythonw & q & " " & q & launcher & q)
        TryOpenProgressPage
        WScript.Quit
    End If
    If fso.FileExists(python) Then
        RunAsync WrapLauncherCmd(q & python & q & " " & q & launcher & q)
        TryOpenProgressPage
        WScript.Quit
    End If

    pythonw = programFiles & "\Python" & tag & "\pythonw.exe"
    python = programFiles & "\Python" & tag & "\python.exe"
    If fso.FileExists(pythonw) Then
        RunAsync WrapLauncherCmd(q & pythonw & q & " " & q & launcher & q)
        TryOpenProgressPage
        WScript.Quit
    End If
    If fso.FileExists(python) Then
        RunAsync WrapLauncherCmd(q & python & q & " " & q & launcher & q)
        TryOpenProgressPage
        WScript.Quit
    End If
End Sub

Function DetectArchSuffix()
    Dim arch, arch2
    arch = WshShell.Environment("Process")("PROCESSOR_ARCHITECTURE")
    arch2 = WshShell.Environment("Process")("PROCESSOR_ARCHITEW6432")
    If UCase(arch) = "X86" And arch2 <> "" Then
        arch = arch2
    End If
    If UCase(arch) = "ARM64" Then
        DetectArchSuffix = "arm64"
    Else
        DetectArchSuffix = "amd64"
    End If
End Function

Function DownloadAndInstallPython(version)
    On Error Resume Next
    Dim suffix, url, tempDir, outPath, ps, cmd, rc
    suffix = DetectArchSuffix()
    url = "https://www.python.org/ftp/python/" & version & "/python-" & version & "-" & suffix & ".exe"
    tempDir = WshShell.ExpandEnvironmentStrings("%TEMP%")
    outPath = tempDir & "\answer_sheet_studio_python_" & version & "_" & suffix & ".exe"

    ps = "$ErrorActionPreference='Stop';" & _
         "$ProgressPreference='SilentlyContinue';" & _
         "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;" & _
         "$url='" & Replace(url, "'", "''") & "';" & _
         "$out='" & Replace(outPath, "'", "''") & "';" & _
         "try { Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing } catch { (New-Object Net.WebClient).DownloadFile($url, $out) };" & _
         "try { $sig=Get-AuthenticodeSignature -FilePath $out; if ($sig.Status -ne 'Valid') { throw ('Invalid installer signature: ' + $sig.Status) } } catch { throw };" & _
         "Start-Process -FilePath $out -ArgumentList '/passive','InstallAllUsers=0','PrependPath=1','Include_pip=1','Include_launcher=1','Include_test=0' -Wait;"

    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command " & q & ps & q
    rc = WshShell.Run(cmd, 0, True)
    DownloadAndInstallPython = (rc = 0)
End Function

probe = q & "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)" & q

' Prefer supported Python versions (3.10+) via the Windows py launcher (pyw) if available.
If CanRun("pyw -3.10 -c " & probe) Then
    RunAsync WrapLauncherCmd("pyw -3.10 " & q & launcher & q)
    TryOpenProgressPage
    WScript.Quit
End If
If CanRun("pyw -3.11 -c " & probe) Then
    RunAsync WrapLauncherCmd("pyw -3.11 " & q & launcher & q)
    TryOpenProgressPage
    WScript.Quit
End If
If CanRun("pyw -3.12 -c " & probe) Then
    RunAsync WrapLauncherCmd("pyw -3.12 " & q & launcher & q)
    TryOpenProgressPage
    WScript.Quit
End If
If CanRun("pyw -3.13 -c " & probe) Then
    RunAsync WrapLauncherCmd("pyw -3.13 " & q & launcher & q)
    TryOpenProgressPage
    WScript.Quit
End If

' Fallback: pythonw (no console window)
If CanRun("pythonw -c " & probe) Then
    RunAsync WrapLauncherCmd("pythonw " & q & launcher & q)
    TryOpenProgressPage
    WScript.Quit
End If

TryRunFromKnownInstall "310"
TryRunFromKnownInstall "311"
TryRunFromKnownInstall "312"
TryRunFromKnownInstall "313"

resp = WshShell.Popup("Python 3.10+ was not found." & vbCrLf & vbCrLf & _
    "Do you want Answer Sheet Studio to download and install Python " & recommendedPython & " automatically now?" & vbCrLf & _
    "(Official installer from python.org)", 0, "Answer Sheet Studio", 36)

If resp = 6 Then ' Yes
    If DownloadAndInstallPython(recommendedPython) Then
        tag = MajorMinorTag(recommendedPython)
        If tag <> "" Then
            TryRunFromKnownInstall tag
        End If
        WshShell.Popup "Python was installed, but could not be detected. Please run Answer Sheet Studio again.", 0, "Answer Sheet Studio", 48
        WScript.Quit
    End If
    WshShell.Popup "Failed to download or install Python automatically." & vbCrLf & vbCrLf & "Please install Python 3.11 (recommended) from python.org, then run Answer Sheet Studio again.", 0, "Answer Sheet Studio", 48
End If

On Error Resume Next
WshShell.Run "https://www.python.org/downloads/windows/", 1, False
WScript.Quit 1
