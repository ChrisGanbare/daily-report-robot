@echo off
:: =============================================================
:: 智能油库日报机器人 —— Windows 一键部署脚本
:: 使用 NSSM 将程序注册为 Windows 系统服务（开机自启、崩溃自动重启）
:: 下载 NSSM：https://nssm.cc/download  解压后将 nssm.exe 放到本目录
:: 用法：以管理员身份运行此脚本
:: =============================================================
setlocal

set SERVICE=DailyReportRobot
set INSTALL_DIR=C:\daily_report_robot
set PYTHON=python

echo [1/5] 创建安装目录 %INSTALL_DIR%
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

echo [2/5] 复制程序文件
copy /Y main.py          "%INSTALL_DIR%\"
copy /Y requirements.txt "%INSTALL_DIR%\"
if exist device_config.xlsx copy /Y device_config.xlsx "%INSTALL_DIR%\"

echo [3/5] 安装 Python 依赖
pip install -r "%INSTALL_DIR%\requirements.txt"

echo [4/5] 注册 Windows 服务（需要 nssm.exe 在当前目录）
if not exist nssm.exe (
    echo 错误：未找到 nssm.exe，请从 https://nssm.cc/download 下载后放到本目录再运行
    pause & exit /b 1
)
:: 若服务已存在则先删除
nssm.exe stop    %SERVICE% 2>nul
nssm.exe remove  %SERVICE% confirm 2>nul

nssm.exe install %SERVICE% %PYTHON% "%INSTALL_DIR%\main.py"
nssm.exe set     %SERVICE% AppDirectory  "%INSTALL_DIR%"
nssm.exe set     %SERVICE% DisplayName   "智能油库日报机器人"
nssm.exe set     %SERVICE% Description   "每日自动拉取油库数据并发送报表"
nssm.exe set     %SERVICE% Start         SERVICE_AUTO_START

:: 崩溃后 30 秒自动重启
nssm.exe set     %SERVICE% AppRestartDelay 30000

:: 标准输出/错误重定向到日志（可选，程序自身已写日志文件）
:: nssm.exe set %SERVICE% AppStdout "%INSTALL_DIR%\stdout.log"
:: nssm.exe set %SERVICE% AppStderr "%INSTALL_DIR%\stderr.log"

echo [5/5] 启动服务
nssm.exe start %SERVICE%

echo.
echo ==============================
echo   部署完成！
echo ==============================
echo   查看状态：sc query %SERVICE%
echo   停止服务：nssm.exe stop %SERVICE%
echo   重启服务：nssm.exe restart %SERVICE%
echo   卸载服务：nssm.exe remove %SERVICE% confirm
echo ==============================
pause
