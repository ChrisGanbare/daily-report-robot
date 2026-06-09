@echo off
:: =============================================================
:: Daily Report Robot - Windows One-Click Deployment Script
:: Uses NSSM to register the program as a Windows service.
:: Download NSSM: https://nssm.cc/download and place nssm.exe here.
:: Usage: Run this script as an administrator.
:: =============================================================
setlocal

set SERVICE=DailyReportRobot
set INSTALL_DIR=C:\daily_report_robot

:: --- Auto-detect Python executable path ---
set PYTHON=
where python >nul 2>&1 && set PYTHON=python
if "%PYTHON%"=="" (
    where py >nul 2>&1 && set PYTHON=py -3
)
if "%PYTHON%"=="" (
    echo Error: Python not found. Please install Python 3.8+ and add it to PATH.
    pause & exit /b 1
)
echo Found Python: %PYTHON%

echo [1/5] Creating installation directory %INSTALL_DIR%
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

echo [2/5] Copying program files
copy /Y main.py          "%INSTALL_DIR%\"
copy /Y requirements.txt "%INSTALL_DIR%\"
if exist device_config.xlsx copy /Y device_config.xlsx "%INSTALL_DIR%\"

echo [3/5] Installing Python dependencies
%PYTHON% -m pip install -r "%INSTALL_DIR%\requirements.txt"
if errorlevel 1 (
    echo Error: Failed to install dependencies. Check network or pip configuration.
    pause & exit /b 1
)

echo [4/5] Registering Windows service (requires nssm.exe in current directory)
if not exist nssm.exe (
    echo Error: nssm.exe not found. Please download it from https://nssm.cc/download and place it here.
    pause & exit /b 1
)

:: Stop and remove the service if it already exists
nssm.exe stop    %SERVICE% 2>nul
nssm.exe remove  %SERVICE% confirm 2>nul

:: Install the service
for /f "delims=" %%i in ('where %PYTHON%') do set PYTHON_EXE=%%i
nssm.exe install %SERVICE% "%PYTHON_EXE%"
nssm.exe set     %SERVICE% AppParameters    "%INSTALL_DIR%\main.py"
nssm.exe set     %SERVICE% AppDirectory     "%INSTALL_DIR%"
nssm.exe set     %SERVICE% DisplayName      "Daily Report Robot"
nssm.exe set     %SERVICE% Description      "Fetches data and sends daily reports automatically."
nssm.exe set     %SERVICE% Start            SERVICE_AUTO_START

:: Restart the service 30 seconds after a crash
nssm.exe set     %SERVICE% AppRestartDelay  30000

:: --- Inject credentials via environment variables ---
:: IMPORTANT: nssm AppEnvironmentExtra must be set in ONE single call.
::            Multiple separate calls OVERWRITE each other - only the last one takes effect.
:: Fill in the real values below and uncomment the entire block.
::
:: nssm.exe set %SERVICE% AppEnvironmentExtra ^
::     "DB_HOST=127.0.0.1" ^
::     "DB_PORT=3306" ^
::     "DB_USER=your_db_user" ^
::     "DB_PASSWORD=your_db_password" ^
::     "DB_NAME=oil" ^
::     "WEBHOOK_URL=https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=YOUR_KEY" ^
::     "FEISHU_APP_ID=cli_xxx" ^
::     "FEISHU_APP_SECRET=xxx" ^
::     "CONFIG_EXCEL_URL=https://your_feishu_or_tencent_doc_url" ^
::     "FEISHU_ALERT_WEBHOOK=https://open.feishu.cn/open-apis/bot/v2/hook/xxx"

:: Redirect stdout/stderr to logs (optional)
:: nssm.exe set %SERVICE% AppStdout "%INSTALL_DIR%\stdout.log"
:: nssm.exe set %SERVICE% AppStderr "%INSTALL_DIR%\stderr.log"

echo [5/5] Starting the service
nssm.exe start %SERVICE%
if errorlevel 1 (
    echo Warning: Failed to start the service. Please run "nssm.exe edit %SERVICE%" to check the configuration.
    pause & exit /b 1
)

echo.
echo ==============================
echo   Deployment complete!
echo ==============================
echo   Check status: sc query %SERVICE%
echo   Edit service: nssm.exe edit %SERVICE%
echo   Stop service: nssm.exe stop %SERVICE%
echo   Restart service: nssm.exe restart %SERVICE%
echo   Remove service: nssm.exe remove %SERVICE% confirm
echo ==============================
pause
endlocal
pause
