@echo off
REM trajectory_splitter GUI 一键启动
REM 双击本文件即可弹出 Tkinter 窗口
REM
REM 依赖:
REM   - e:\DuDao\project_weixiu\.venv\Scripts\python.exe
REM   - xlrd / xlwt / pyyaml / tkinter (内置)
REM
REM 如果虚拟环境被移动, 改下面 VENV_PYTHON 这一行。

set VENV_PYTHON=E:\DuDao\project_weixiu\.venv\Scripts\python.exe
set PROJECT_ROOT=E:\DuDao\project_weixiu

cd /d "%PROJECT_ROOT%"
"%VENV_PYTHON%" -m trajectory_splitter
if errorlevel 1 (
    echo.
    echo [trajectory_splitter] 启动失败, 错误码 %errorlevel%
    echo 检查 venv 路径是否正确, 或运行 .venv\Scripts\pip install -r trajectory_splitter\requirements.txt
    pause
)