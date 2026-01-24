Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
repo = fso.GetParentFolderName(WScript.ScriptFullName)

q = Chr(34)
launcher = repo & "\launcher_headless.py"
recommendedPython = WshShell.Environment("Process")("ANSWER_SHEET_PYTHON_VERSION")
If recommendedPython = "" Then
    recommendedPython = "3.11.8"
End If

Function DetectLatestRWindowsExe()
    On Error Resume Next
    Dim ps, cmd, rc, out, tempDir, outPath, f
    tempDir = WshShell.ExpandEnvironmentStrings("%TEMP%")
    outPath = tempDir & "\answer_sheet_studio_r_latest.txt"
    ' Query CRAN and extract the latest R-x.y.z-win.exe filename
    ps = "$ErrorActionPreference='Stop';" & _
         "$base='https://cran.r-project.org/bin/windows/base/';" & _
         "$html=(Invoke-WebRequest -Uri $base -UseBasicParsing).Content;" & _
         "$ms=[regex]::Matches($html,'R-([0-9]+\\.[0-9]+\\.[0-9]+)-win\\.exe');" & _
         "$vers=@(); foreach($m in $ms){ $vers += [version]$m.Groups[1].Value };" & _
         "if($vers.Count -eq 0){ throw 'No R installer found' };" & _
         "$v=($vers | Sort-Object -Descending | Select-Object -First 1);" & _
         "('R-' + $v.ToString() + '-win.exe') | Set-Content -Encoding ASCII -NoNewline '" & Replace(outPath, "\", "\\") & "';"
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " & q & ps & q
    rc = WshShell.Run(cmd, 0, True)
    If rc <> 0 Then
        DetectLatestRWindowsExe = ""
        Exit Function
    End If
    If Not fso.FileExists(outPath) Then
        DetectLatestRWindowsExe = ""
        Exit Function
    End If
    Set f = fso.OpenTextFile(outPath, 1, False)
    out = f.ReadAll
    f.Close
    DetectLatestRWindowsExe = Trim(out)
End Function

Function NormalizeEnvPath(value)
    On Error Resume Next
    Dim v
    v = Trim(CStr(value))
    ' If the env var is undefined, ExpandEnvironmentStrings returns "%NAME%".
    If InStr(v, "%") > 0 Then
        v = ""
    End If
    NormalizeEnvPath = v
End Function

Function VersionGreater(a, b)
    On Error Resume Next
    Dim aParts, bParts, i, ai, bi
    aParts = Split(CStr(a), ".")
    bParts = Split(CStr(b), ".")
    For i = 0 To 2
        ai = 0
        bi = 0
        If i <= UBound(aParts) Then ai = CLng(Val(aParts(i)))
        If i <= UBound(bParts) Then bi = CLng(Val(bParts(i)))
        If ai > bi Then
            VersionGreater = True
            Exit Function
        End If
        If ai < bi Then
            VersionGreater = False
            Exit Function
        End If
    Next
    VersionGreater = False
End Function

Function RegReadString(key)
    On Error Resume Next
    Dim v
    v = ""
    v = WshShell.RegRead(key)
    If Err.Number <> 0 Then
        Err.Clear
        v = ""
    End If
    RegReadString = Trim(CStr(v))
End Function

Function FindRscriptFromInstallPath(installPath)
    On Error Resume Next
    Dim base, candidate
    base = Trim(CStr(installPath))
    If base = "" Then
        FindRscriptFromInstallPath = ""
        Exit Function
    End If
    If Right(base, 1) = "\" Then
        base = Left(base, Len(base) - 1)
    End If
    candidate = base & "\bin\Rscript.exe"
    If Not fso.FileExists(candidate) Then
        candidate = base & "\bin\x64\Rscript.exe"
    End If
    If Not fso.FileExists(candidate) Then
        candidate = base & "\bin\i386\Rscript.exe"
    End If
    If fso.FileExists(candidate) Then
        FindRscriptFromInstallPath = candidate
        Exit Function
    End If
    FindRscriptFromInstallPath = ""
End Function

Function FindRscriptFromRegistry()
    On Error Resume Next
    Dim keys, key, installPath, cand
    keys = Array( _
        "HKLM\SOFTWARE\R-core\R\InstallPath", _
        "HKLM\SOFTWARE\R-core\R64\InstallPath", _
        "HKLM\SOFTWARE\WOW6432Node\R-core\R\InstallPath", _
        "HKLM\SOFTWARE\WOW6432Node\R-core\R64\InstallPath", _
        "HKCU\SOFTWARE\R-core\R\InstallPath", _
        "HKCU\SOFTWARE\R-core\R64\InstallPath", _
        "HKCU\SOFTWARE\WOW6432Node\R-core\R\InstallPath", _
        "HKCU\SOFTWARE\WOW6432Node\R-core\R64\InstallPath" _
    )
    For Each key In keys
        installPath = RegReadString(key)
        cand = FindRscriptFromInstallPath(installPath)
        If cand <> "" Then
            FindRscriptFromRegistry = cand
            Exit Function
        End If
    Next
    FindRscriptFromRegistry = ""
End Function

Function FindRscriptExe()
    On Error Resume Next
    Dim localAppData, programFiles, programFilesX86, programW6432, roots, root, rRoot, folder, subfolder
    Dim localPrograms
    Dim name, ver, bestVer, bestPath, candidate

    bestVer = ""
    bestPath = ""

    localAppData = NormalizeEnvPath(WshShell.ExpandEnvironmentStrings("%LOCALAPPDATA%"))
    programFiles = NormalizeEnvPath(WshShell.ExpandEnvironmentStrings("%ProgramFiles%"))
    programFilesX86 = NormalizeEnvPath(WshShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%"))
    programW6432 = NormalizeEnvPath(WshShell.ExpandEnvironmentStrings("%ProgramW6432%"))

    localPrograms = ""
    If localAppData <> "" Then
        localPrograms = localAppData & "\Programs"
    End If

    roots = Array(localPrograms, programW6432, programFiles, programFilesX86)

    For Each root In roots
        If root <> "" Then
            rRoot = root & "\R"
            If fso.FolderExists(rRoot) Then
                Set folder = fso.GetFolder(rRoot)
                For Each subfolder In folder.SubFolders
                    name = subfolder.Name
                    If LCase(Left(name, 2)) = "r-" Then
                        ver = Mid(name, 3)

                        candidate = subfolder.Path & "\bin\Rscript.exe"
                        If Not fso.FileExists(candidate) Then
                            candidate = subfolder.Path & "\bin\x64\Rscript.exe"
                        End If
                        If Not fso.FileExists(candidate) Then
                            candidate = subfolder.Path & "\bin\i386\Rscript.exe"
                        End If

                        If fso.FileExists(candidate) Then
                            If bestVer = "" Or VersionGreater(ver, bestVer) Then
                                bestVer = ver
                                bestPath = candidate
                            End If
                        End If
                    End If
                Next
            End If
        End If
    Next

    If bestPath = "" Then
        bestPath = FindRscriptFromRegistry()
    End If

    FindRscriptExe = bestPath
End Function

Function DownloadAndInstallR()
    On Error Resume Next
    Dim exeName, url, tempDir, outPath, ps, cmd, rc
    exeName = DetectLatestRWindowsExe()
    If exeName = "" Then
        DownloadAndInstallR = False
        Exit Function
    End If
    url = "https://cran.r-project.org/bin/windows/base/" & exeName
    tempDir = WshShell.ExpandEnvironmentStrings("%TEMP%")
    outPath = tempDir & "\answer_sheet_studio_" & exeName

    ps = "$ErrorActionPreference='Stop';" & _
         "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;" & _
         "$url='" & Replace(url, "'", "''") & "';" & _
         "$out='" & Replace(outPath, "'", "''") & "';" & _
         "try { Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing } catch { (New-Object Net.WebClient).DownloadFile($url, $out) };" & _
         "try { $sig=Get-AuthenticodeSignature -FilePath $out; if ($sig.Status -ne 'Valid') { throw ('Invalid installer signature: ' + $sig.Status) } } catch { throw };" & _
         "Start-Process -FilePath $out;"

    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " & q & ps & q
    rc = WshShell.Run(cmd, 0, True)
    DownloadAndInstallR = (rc = 0)
End Function

Sub EnsureRInstalledOrExit()
    On Error Resume Next
    If CanRun("Rscript -e " & q & "quit(status=0)" & q) Then
        Exit Sub
    End If

    Dim rscriptExe
    rscriptExe = FindRscriptExe()
    If rscriptExe <> "" Then
        If CanRun(q & rscriptExe & q & " -e " & q & "quit(status=0)" & q) Then
            Exit Sub
        End If
    End If
    ' R is optional; skip installing unless explicitly enabled.
    Dim installR
    installR = LCase(Trim(WshShell.Environment("Process")("ANSWER_SHEET_INSTALL_R")))
    If installR <> "1" And installR <> "true" And installR <> "yes" Then
        Exit Sub
    End If

    ' Optional: download and start installing R from CRAN, but do not block startup.
    If DownloadAndInstallR() Then
        WshShell.Popup "R installer started (optional)." & vbCrLf & vbCrLf & _
            "Answer Sheet Studio will continue without R. Restart after installing R to enable ggplot2 plots.", 0, "Answer Sheet Studio", 64
    Else
        WshShell.Popup "R was not found." & vbCrLf & vbCrLf & _
            "Answer Sheet Studio will continue without R. Install R from CRAN to enable ggplot2 plots.", 0, "Answer Sheet Studio", 48
        On Error Resume Next
        WshShell.Run "https://cran.r-project.org/bin/windows/base/", 1, False
    End If
End Sub

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
        EnsureRInstalledOrExit
        RunAsync q & pythonw & q & " " & q & launcher & q
        WScript.Quit
    End If
    If fso.FileExists(python) Then
        EnsureRInstalledOrExit
        RunAsync q & python & q & " " & q & launcher & q
        WScript.Quit
    End If

    pythonw = programFiles & "\Python" & tag & "\pythonw.exe"
    python = programFiles & "\Python" & tag & "\python.exe"
    If fso.FileExists(pythonw) Then
        EnsureRInstalledOrExit
        RunAsync q & pythonw & q & " " & q & launcher & q
        WScript.Quit
    End If
    If fso.FileExists(python) Then
        EnsureRInstalledOrExit
        RunAsync q & python & q & " " & q & launcher & q
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

    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " & q & ps & q
    rc = WshShell.Run(cmd, 0, True)
    DownloadAndInstallPython = (rc = 0)
End Function

probe = q & "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)" & q

' Prefer supported Python versions (3.10+) via the Windows py launcher (pyw) if available.
If CanRun("pyw -3.10 -c " & probe) Then
    EnsureRInstalledOrExit
    RunAsync "pyw -3.10 " & q & launcher & q
    WScript.Quit
End If
If CanRun("pyw -3.11 -c " & probe) Then
    EnsureRInstalledOrExit
    RunAsync "pyw -3.11 " & q & launcher & q
    WScript.Quit
End If
If CanRun("pyw -3.12 -c " & probe) Then
    EnsureRInstalledOrExit
    RunAsync "pyw -3.12 " & q & launcher & q
    WScript.Quit
End If
If CanRun("pyw -3.13 -c " & probe) Then
    EnsureRInstalledOrExit
    RunAsync "pyw -3.13 " & q & launcher & q
    WScript.Quit
End If

' Fallback: pythonw (no console window)
If CanRun("pythonw -c " & probe) Then
    EnsureRInstalledOrExit
    RunAsync "pythonw " & q & launcher & q
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
