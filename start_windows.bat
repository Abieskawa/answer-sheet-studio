@echo off
setlocal

set "REPO_DIR=%~dp0"
set "REPO_DIR=%REPO_DIR:~0,-1%"

if not exist "%REPO_DIR%\outputs" mkdir "%REPO_DIR%\outputs" >nul 2>&1
set "LOG=%REPO_DIR%\outputs\start_windows_bat.log"
echo [%date% %time%] start_windows.bat launched, repo=%REPO_DIR%>>"%LOG%"

set "PROBE=import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)"
set "ANSWER_SHEET_PROGRESS=cli"
set "ANSWER_SHEET_OPEN_BROWSER=0"
set "ANSWER_SHEET_WIZARD=1"
if "%ANSWER_SHEET_PYTHON_VERSION%"=="" set "ANSWER_SHEET_PYTHON_VERSION=3.11.8"

where Rscript >nul 2>&1
if errorlevel 1 (
  echo [%date% %time%] Rscript not found>>"%LOG%"
  echo R was not found. R is optional (for nicer ggplot2 plots).
  choice /c YN /m "Install R now? (optional)"
  if errorlevel 2 goto :PYTHON_LAUNCH

  echo [%date% %time%] starting R download/install>>"%LOG%"
  powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ErrorActionPreference='Stop';" ^
    "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;" ^
    "$base='https://cran.r-project.org/bin/windows/base/';" ^
    "$html=(Invoke-WebRequest -Uri $base -UseBasicParsing).Content;" ^
    "$ms=[regex]::Matches($html,'R-([0-9]+\\.[0-9]+\\.[0-9]+)-win\\.exe');" ^
    "$vers=@(); foreach($m in $ms){ $vers += [version]$m.Groups[1].Value };" ^
    "if($vers.Count -eq 0){ throw 'No R installer found' };" ^
    "$v=($vers | Sort-Object -Descending | Select-Object -First 1).ToString();" ^
    "$exe=('R-{0}-win.exe' -f $v);" ^
    "$url=($base + $exe);" ^
    "$out=(Join-Path $env:TEMP ('answer_sheet_studio_' + $exe));" ^
    "$ProgressPreference='SilentlyContinue';" ^
    "try { Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing } catch { (New-Object Net.WebClient).DownloadFile($url, $out) };" ^
    "$sig=Get-AuthenticodeSignature -FilePath $out; if($sig.Status -ne 'Valid'){ throw ('Invalid installer signature: ' + $sig.Status) };" ^
    "Start-Process -FilePath $out"

  echo [%date% %time%] R installer started>>"%LOG%"
)

:PYTHON_LAUNCH
for %%P in (pyw -3.11 pyw -3.10 pyw -3.12 pyw -3.13) do (
  %%P -c "%PROBE%" >nul 2>&1
  if not errorlevel 1 (
    echo [%date% %time%] using %%P>>"%LOG%"
    echo [%date% %time%] running in terminal: %%P>>"%LOG%"
    %%P "%REPO_DIR%\launcher_headless.py"
    goto :END
  )
)

for %%P in (pythonw python) do (
  %%P -c "%PROBE%" >nul 2>&1
  if not errorlevel 1 (
    echo [%date% %time%] using %%P>>"%LOG%"
    echo [%date% %time%] running in terminal: %%P>>"%LOG%"
    %%P "%REPO_DIR%\launcher_headless.py"
    goto :END
  )
)

echo [%date% %time%] Python 3.10+ not found or not runnable>>"%LOG%"
echo Python 3.10+ was not found.
echo We can download and install Python %ANSWER_SHEET_PYTHON_VERSION% automatically (official installer from python.org).
choice /c YN /m "Download and install now?"
if errorlevel 2 exit /b 1

echo [%date% %time%] starting Python download/install %ANSWER_SHEET_PYTHON_VERSION%>>"%LOG%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop';" ^
  "$v=$env:ANSWER_SHEET_PYTHON_VERSION; if(-not $v){$v='3.11.8'};" ^
  "$arch=$env:PROCESSOR_ARCHITECTURE; $arch2=$env:PROCESSOR_ARCHITEW6432; if($arch -eq 'x86' -and $arch2){$arch=$arch2};" ^
  "$suffix=if($arch -eq 'ARM64'){'arm64'}else{'amd64'};" ^
  "$url=('https://www.python.org/ftp/python/{0}/python-{0}-{1}.exe' -f $v,$suffix);" ^
  "$out=(Join-Path $env:TEMP ('answer_sheet_studio_python_{0}_{1}.exe' -f $v,$suffix));" ^
  "$ProgressPreference='SilentlyContinue';" ^
  "try { Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing } catch { (New-Object Net.WebClient).DownloadFile($url, $out) };" ^
  "$sig=Get-AuthenticodeSignature -FilePath $out; if($sig.Status -ne 'Valid'){ throw ('Invalid installer signature: ' + $sig.Status) };" ^
  "Start-Process -FilePath $out -ArgumentList '/passive','InstallAllUsers=0','PrependPath=1','Include_pip=1','Include_launcher=1','Include_test=0' -Wait;"

echo [%date% %time%] Python installer finished; retrying>>"%LOG%"
goto :RETRY

:RETRY
for %%P in (pyw -3.11 pyw -3.10 pyw -3.12 pyw -3.13) do (
  %%P -c "%PROBE%" >nul 2>&1
  if not errorlevel 1 (
    echo [%date% %time%] using %%P>>"%LOG%"
    echo [%date% %time%] running in terminal: %%P>>"%LOG%"
    %%P "%REPO_DIR%\launcher_headless.py"
    goto :END
  )
)
for %%P in (pythonw python) do (
  %%P -c "%PROBE%" >nul 2>&1
  if not errorlevel 1 (
    echo [%date% %time%] using %%P>>"%LOG%"
    echo [%date% %time%] running in terminal: %%P>>"%LOG%"
    %%P "%REPO_DIR%\launcher_headless.py"
    goto :END
  )
)
echo [%date% %time%] Python still not detected after install>>"%LOG%"
echo Python still not detected. Please restart this launcher, or log out/in, then try again.
exit /b 1

:END
echo [%date% %time%] finished>>"%LOG%"
exit /b 0
