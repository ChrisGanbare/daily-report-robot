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

:: ── 自动检测 Python 可执行路径 ──────────────────────────────
set PYTHON=
where python >nul 2>&1 && set PYTHON=python
if "%PYTHON%"=="" (
    where py >nul 2>&1 && set PYTHON=py -3
)
if "%PYTHON%"=="" (
    echo 错误：未找到 Python，请先安装 Python 3.8+ 并加入 PATH
    pause & exit /b 1
)
echo 检测到 Python：%PYTHON%

echo [1/5] 创建安装目录 %INSTALL_DIR%
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

echo [2/5] 复制程序文件
copy /Y main.py          "%INSTALL_DIR%\"
copy /Y requirements.txt "%INSTALL_DIR%\"
if exist device_config.xlsx copy /Y device_config.xlsx "%INSTALL_DIR%\"

echo [3/5] 安装 Python 依赖
%PYTHON% -m pip install -r "%INSTALL_DIR%\requirements.txt"
if errorlevel 1 (
    echo 错误：依赖安装失败，请检查网络或 pip 配置
    pause & exit /b 1
)

echo [4/5] 注册 Windows 服务（需要 nssm.exe 在当前目录）
if not exist nssm.exe (
    echo 错误：未找到 nssm.exe，请从 https://nssm.cc/download 下载后放到本目录再运行
    pause & exit /b 1
)

:: 若服务已存在则先停止并删除
nssm.exe stop    %SERVICE% 2>nul
nssm.exe remove  %SERVICE% confirm 2>nul

:: 注册服务（仅指定可执行程序，参数通过 AppParameters 单独设置）
for /f "delims=" %%i in ('where %PYTHON%') do set PYTHON_EXE=%%i
nssm.exe install %SERVICE% "%PYTHON_EXE%"
nssm.exe set     %SERVICE% AppParameters    "%INSTALL_DIR%\main.py"
nssm.exe set     %SERVICE% AppDirectory     "%INSTALL_DIR%"
nssm.exe set     %SERVICE% DisplayName      "智能油库日报机器人"
nssm.exe set     %SERVICE% Description      "每日自动拉取油库数据并发送报表"
nssm.exe set     %SERVICE% Start            SERVICE_AUTO_START

:: 崩溃后 30 秒自动重启
nssm.exe set     %SERVICE% AppRestartDelay  30000

:: ── 敏感凭据建议通过环境变量注入，避免写在代码中 ──
:: 取消注释并填入真实值后重新运行脚本，或在服务注册后通过 nssm.exe edit 配置
:: nssm.exe set %SERVICE% AppEnvironmentExtra "DB_HOST=127.0.0.1"
:: nssm.exe set %SERVICE% AppEnvironmentExtra "DB_PASSWORD=your_password"
:: nssm.exe set %SERVICE% AppEnvironmentExtra "WEBHOOK_URL=https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=YOUR_KEY"
:: nssm.exe set %SERVICE% AppEnvironmentExtra "FEISHU_APP_ID=cli_xxx"
:: nssm.exe set %SERVICE% AppEnvironmentExtra "FEISHU_APP_SECRET=xxx"
:: nssm.exe set %SERVICE% AppEnvironmentExtra "FEISHU_ALERT_WEBHOOK=https://open.feishu.cn/open-apis/bot/v2/hook/xxx"

:: 标准输出/错误重定向到日志（可选，程序自身已写日志文件）
:: nssm.exe set %SERVICE% AppStdout "%INSTALL_DIR%\stdout.log"
:: nssm.exe set %SERVICE% AppStderr "%INSTALL_DIR%\stderr.log"

echo [5/5] 启动服务
nssm.exe start %SERVICE%
if errorlevel 1 (
    echo 警告：服务启动失败，请运行 "nssm.exe edit %SERVICE%" 检查配置
    pause & exit /b 1
)

echo.
echo ==============================
echo   部署完成！
echo ==============================
echo   查看状态：sc query %SERVICE%
echo   图形配置：nssm.exe edit %SERVICE%
echo   停止服务：nssm.exe stop %SERVICE%
echo   重启服务：nssm.exe restart %SERVICE%
echo   卸载服务：nssm.exe remove %SERVICE% confirm
echo ==============================
pause
