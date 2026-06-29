@echo off
:: ============================================================================
:: CoA Lookup (BD / Miltenyi / BioLegend) - Windows Setup & Run
:: Double-click to set up Python, install packages, and launch the app.
:: Auto-installs Python 3.12 (per-user, no admin needed) if not already present.
:: ============================================================================

setlocal EnableDelayedExpansion
cd /d "%~dp0"

set LOG=%~dp0setup_run.log
> "%LOG%" echo CoA Lookup - Setup Log - %date% %time%
>> "%LOG%" echo Working dir: %CD%
>> "%LOG%" echo.

echo.
echo  ============================================
echo   CoA Lookup ^(BD / Miltenyi / BioLegend^)
echo  ============================================
echo   Log: %LOG%
echo.

:: ── Step 1: locate or install Python 3.9+ ───────────────────────────────────
echo [1/5] Checking Python...
>> "%LOG%" echo [STEP 1] Looking for Python 3.9+
set PYTHON=

:: 1a. Try the "py" launcher first - most reliable on Windows
where py >nul 2>&1
if not errorlevel 1 (
    py -3 -c "import sys; sys.exit(0 if sys.version_info>=(3,9) else 1)" >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON=py -3"
        >> "%LOG%" echo Found via: py -3
        goto :have_python
    )
)

:: 1b. Try "python" - but reject the Microsoft Store stub.
::     The stub passes `--version` but its full path is in WindowsApps.
where python >nul 2>&1
if not errorlevel 1 (
    for /f "delims=" %%P in ('where python 2^>nul') do (
        echo %%P | findstr /I "WindowsApps" >nul
        if errorlevel 1 (
            "%%P" -c "import sys; sys.exit(0 if sys.version_info>=(3,9) else 1)" >nul 2>&1
            if not errorlevel 1 (
                set "PYTHON=%%P"
                >> "%LOG%" echo Found via: %%P
                goto :have_python
            )
        ) else (
            >> "%LOG%" echo Skipping Microsoft Store stub: %%P
        )
    )
)

:: 1c. Try python3
where python3 >nul 2>&1
if not errorlevel 1 (
    python3 -c "import sys; sys.exit(0 if sys.version_info>=(3,9) else 1)" >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON=python3"
        >> "%LOG%" echo Found via: python3
        goto :have_python
    )
)

:: ── No Python found: try to install it automatically ──────────────────────
echo.
echo  Python 3.9+ not found. Attempting auto-install...
echo  ^(This may take 1-3 minutes. No admin required.^)
echo.
>> "%LOG%" echo Python not found - trying auto-install

:: Try winget (built into Windows 10 1809+ / Windows 11)
where winget >nul 2>&1
if not errorlevel 1 (
    echo  Installing Python 3.12 via winget...
    >> "%LOG%" echo Trying winget install
    winget install -e --id Python.Python.3.12 --silent --scope user --accept-source-agreements --accept-package-agreements >> "%LOG%" 2>&1
    if not errorlevel 1 (
        >> "%LOG%" echo winget install succeeded - refreshing PATH
        call :refresh_path
        :: Retry detection
        where py >nul 2>&1
        if not errorlevel 1 (
            py -3 -c "import sys" >nul 2>&1
            if not errorlevel 1 (
                set "PYTHON=py -3"
                echo  [OK]  Python installed via winget.
                >> "%LOG%" echo OK - py launcher works after winget
                goto :have_python
            )
        )
    ) else (
        echo  winget install failed. Trying alternate method...
        >> "%LOG%" echo winget install failed
    )
)

:: PowerShell fallback: download python.org installer and run silently per-user
echo  Downloading Python 3.12 installer from python.org...
>> "%LOG%" echo Trying python.org installer download
set "PYINSTALL=%TEMP%\python-3.12.7-amd64.exe"
if exist "!PYINSTALL!" del "!PYINSTALL!" >nul 2>&1

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ProgressPreference='SilentlyContinue'; try { Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe' -OutFile '!PYINSTALL!' -UseBasicParsing } catch { exit 1 }" >> "%LOG%" 2>&1

if not exist "!PYINSTALL!" (
    echo  [ERROR] Could not download Python installer.
    echo          Check your internet connection.
    >> "%LOG%" echo Download failed
    goto :need_manual_install
)

echo  Installing Python ^(per-user, this takes about a minute^)...
>> "%LOG%" echo Running silent installer
"!PYINSTALL!" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_launcher=1 SimpleInstall=1 >> "%LOG%" 2>&1
set INST_RC=!errorlevel!
del "!PYINSTALL!" >nul 2>&1
>> "%LOG%" echo Installer exit code: !INST_RC!

call :refresh_path

:: Re-detect after install
where py >nul 2>&1
if not errorlevel 1 (
    py -3 -c "import sys" >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON=py -3"
        echo  [OK]  Python installed.
        goto :have_python
    )
)
where python >nul 2>&1
if not errorlevel 1 (
    for /f "delims=" %%P in ('where python 2^>nul') do (
        echo %%P | findstr /I "WindowsApps" >nul
        if errorlevel 1 (
            "%%P" -c "import sys" >nul 2>&1
            if not errorlevel 1 (
                set "PYTHON=%%P"
                echo  [OK]  Python installed.
                goto :have_python
            )
        )
    )
)

:need_manual_install
echo.
echo  [ERROR] Could not auto-install Python.
echo.
echo   Please install Python 3.9+ manually from:
echo     https://www.python.org/downloads/
echo   IMPORTANT: tick "Add Python to PATH" during install.
echo   Then double-click this .bat file again.
echo.
>> "%LOG%" echo Manual install required
goto :fail

:have_python
!PYTHON! --version
>> "%LOG%" echo Using PYTHON=!PYTHON!
echo.

:: ── Step 2: virtual environment ─────────────────────────────────────────────
echo [2/5] Setting up virtual environment...
>> "%LOG%" echo [STEP 2] venv
set VENV_DIR=.venv

if not exist "%VENV_DIR%\Scripts\activate.bat" (
    >> "%LOG%" echo Creating .venv...
    !PYTHON! -m venv "%VENV_DIR%" >> "%LOG%" 2>&1
    if errorlevel 1 (
        echo  [ERROR] Failed to create virtual environment.
        echo          Trying without venv ^(global install^)...
        >> "%LOG%" echo venv creation failed - using global Python
        goto :skip_venv
    )
    echo  [OK]  Created .venv
) else (
    echo  [OK]  Found existing .venv
)

call "%VENV_DIR%\Scripts\activate.bat"
set PIPCMD=python -m pip
echo  [OK]  Virtual environment active.
goto :install_packages

:skip_venv
:: Use the system Python directly. Note we still use --user to avoid
:: needing admin rights.
set PIPCMD=!PYTHON! -m pip
set NO_VENV=1
>> "%LOG%" echo Using system Python (no venv)

:install_packages
:: ── Step 3: install packages ────────────────────────────────────────────────
echo.
echo [3/5] Installing Python packages...
>> "%LOG%" echo [STEP 3] pip install

:: Upgrade pip quietly (non-fatal)
%PIPCMD% install --upgrade pip --quiet >> "%LOG%" 2>&1

set PIPFLAGS=
if defined NO_VENV set PIPFLAGS=--user

%PIPCMD% install %PIPFLAGS% flask requests cloudscraper beautifulsoup4 openpyxl pypdf
if errorlevel 1 (
    echo.
    echo  [WARN] First pip install attempt failed. Retrying with trusted-host bypass...
    >> "%LOG%" echo First pip attempt failed - retrying with trusted-host
    %PIPCMD% install %PIPFLAGS% --trusted-host pypi.org --trusted-host files.pythonhosted.org flask requests cloudscraper beautifulsoup4 openpyxl pypdf
    if errorlevel 1 (
        echo  [ERROR] Package installation failed.
        echo          Check internet connection or corporate firewall settings.
        echo          Log: %LOG%
        >> "%LOG%" echo pip install failed twice
        goto :fail
    )
)

:: Try lxml separately - optional, do not fail the whole install if it errors
%PIPCMD% install %PIPFLAGS% lxml --quiet >> "%LOG%" 2>&1
if errorlevel 1 (
    echo  [INFO] lxml unavailable - the app will use the stdlib HTML parser instead.
    >> "%LOG%" echo lxml install failed - falling back to html.parser
)

echo  [OK]  All packages ready.
>> "%LOG%" echo Packages installed

:: ── Step 4: sanity check ────────────────────────────────────────────────────
echo.
echo [4/5] Verifying imports...
>> "%LOG%" echo [STEP 4] Import sanity check
python -c "import flask, requests, cloudscraper, bs4, openpyxl, pypdf; print('  [OK]  flask, requests, cloudscraper, bs4, openpyxl, pypdf all import')"
if errorlevel 1 (
    echo  [ERROR] One of the required packages failed to import.
    echo          See log for details: %LOG%
    >> "%LOG%" echo Import sanity check failed
    goto :fail
)
>> "%LOG%" echo Import sanity check passed

:: ── Step 5: port check + launch ─────────────────────────────────────────────
echo.
echo [5/5] Checking port 5050...
>> "%LOG%" echo [STEP 5] Port check
set PORT=5050
set EXISTING_PID=

netstat -ano > "%TEMP%\blnet.txt" 2>nul
for /f "tokens=1,2,3,4,5" %%A in ('findstr /C:":%PORT% " "%TEMP%\blnet.txt" 2^>nul') do (
    if "%%E" neq "" if "!EXISTING_PID!"=="" call :check_listening "%%D" "%%E"
)
del "%TEMP%\blnet.txt" >nul 2>&1

if not "!EXISTING_PID!"=="" (
    echo.
    echo  [WARN] Port %PORT% is already in use by PID !EXISTING_PID!.
    set /p CONFIRM="         Kill it and start fresh? [y/N]: "
    if /i "!CONFIRM!"=="y" (
        taskkill /PID !EXISTING_PID! /F >nul 2>&1
        echo  [OK]  Killed PID !EXISTING_PID!.
        timeout /t 1 /nobreak >nul
    ) else (
        echo  Opening browser to existing server...
        start "" "http://localhost:%PORT%"
        goto :done
    )
)

:: ── Launch ──────────────────────────────────────────────────────────────────
echo.
echo  ============================================
echo   Server: http://localhost:%PORT%
echo   Keep this window open while using the app.
echo   Press Ctrl+C to stop the server.
echo  ============================================
echo.
>> "%LOG%" echo Launching Flask

start "" /b cmd /c "timeout /t 2 /nobreak >nul && start http://localhost:%PORT%"

python app.py
>> "%LOG%" echo Flask exit code: %errorlevel%

if errorlevel 1 (
    echo.
    echo  [WARN] Server stopped with an error.
    echo         Check the output above for details.
    echo         Log: %LOG%
)
goto :done

:: ── Helpers ─────────────────────────────────────────────────────────────────
:check_listening
echo %~1 | findstr "LISTENING" >nul 2>&1
if errorlevel 1 goto :eof
set EXISTING_PID=%~2
goto :eof

:refresh_path
:: Refresh PATH from the registry (winget/installer changes don't apply
:: to the current cmd session until we manually reload).
for /f "skip=2 tokens=1,2*" %%A in ('reg query "HKCU\Environment" /v PATH 2^>nul') do set "USER_PATH=%%C"
for /f "skip=2 tokens=1,2*" %%A in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PATH 2^>nul') do set "SYS_PATH=%%C"
set "PATH=!USER_PATH!;!SYS_PATH!;!PATH!"
:: Add common per-user Python install paths just in case PATH wasn't updated
set "PATH=%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python311\Scripts;%LOCALAPPDATA%\Programs\Python\Python311;%LOCALAPPDATA%\Programs\Python\Launcher;%PATH%"
goto :eof

:fail
>> "%LOG%" echo *** Setup failed ***
echo.
echo  *** Setup failed. ***
echo.
echo  Log saved to: %LOG%
echo  Share this file if you need help troubleshooting.

:done
echo.
echo  Press any key to close this window...
pause >nul
