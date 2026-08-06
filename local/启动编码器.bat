@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Starting Vigor Encoder LAN server...
echo 本机访问: http://127.0.0.1:8765
echo 局域网地址请在窗口内查看（发给其他同事）
echo 如果其他电脑无法访问，请在 Windows 防火墙中允许 Python 访问网络
echo.

REM 优先使用内置 Python（已安装 openpyxl，支持 CRM 模板导出），其次系统 Python
if exist "C:\Users\Monk Chen\.workbuddy\binaries\python\versions\3.13.12\python.exe" (
    set PY=C:\Users\Monk Chen\.workbuddy\binaries\python\versions\3.13.12\python.exe
) else (
    set PY=python
)

start "" http://127.0.0.1:8765
"%PY%" server.py
pause

