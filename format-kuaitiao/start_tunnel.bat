@echo off
chcp 65001 >nul
title 论文排版 - 公网隧道模式

echo =============================================
echo   论文自动排版 - 公网分享模式
echo =============================================
echo.

:: 启动后端
echo [1/2] 启动后端服务...
start "排版后端" cmd /c "cd /d "%~dp0" && python -m uvicorn server:app --host 0.0.0.0 --port 8000"
timeout /t 2 >nul

:: 启动 SSH 隧道
echo [2/2] 启动公网隧道 (serveo.net)...
echo.
echo 连接成功后，终端会显示公网地址。如：
echo   Forwarding HTTP traffic from [https://xxxxx.serveo.net]
echo.
echo 复制这个地址发给别人即可使用！
echo.
echo 按 Ctrl+C 可停止隧道，关闭后端窗口停止服务。
echo =============================================
echo.

ssh -o StrictHostKeyChecking=accept-new -R 80:localhost:8000 serveo.net
