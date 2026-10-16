@echo off
REM trajectory_splitter GUI launcher
REM Double-click this file to open the Tkinter window.
REM
REM Requires:
REM   - e:\DuDao\project_weixiu\.venv\Scripts\python.exe
REM   - pywin32 (Excel COM) / pyyaml / tkinter (built-in)
REM
REM If the venv has moved, update VENV_PYTHON below.

set VENV_PYTHON=E:\DuDao\project_weixiu\.venv\Scripts\python.exe
set PROJECT_ROOT=E:\DuDao\project_weixiu

cd /d "%PROJECT_ROOT%"
"%VENV_PYTHON%" -m trajectory_splitter
if errorlevel 1 (
    echo.
    echo [trajectory_splitter] launch failed, error code %errorlevel%
    echo check the venv path, or run: .venv\Scripts\pip install -r trajectory_splitter\requirements.txt
    pause
)
