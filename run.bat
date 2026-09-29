@echo off
REM 启动 Excel 指定列生成条形码工具（Windows）
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo 未找到虚拟环境，请先联系开发者。
    pause
    exit /b 1
)
.venv\Scripts\python.exe app.py
pause
