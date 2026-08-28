@echo off
REM trajectory_splitter web GUI launcher (pywebview).
REM Double-click this file to open the HTML/CSS/JS frontend in a native window.
REM The old Tk GUI is still available via launch_gui.bat.
REM
REM Requires:
REM   - e:\DuDao\project_weixiu\.venv\Scripts\python.exe
REM   - pywebview (already installed: 6.2.1)
REM   - xlrd / xlwt / pyyaml / tkinter (built-in)
REM   - Edge WebView2 runtime (built into Win 10/11; on Win 10 older builds,
REM     install the Evergreen Runtime from Microsoft if missing)
REM
REM If the venv has moved, update VENV_PYTHON below.

set VENV_PYTHON=E:\DuDao\project_weixiu\.venv\Scripts\python.exe
set PROJECT_ROOT=E:\DuDao\project_weixiu

cd /d "%PROJECT_ROOT%"
"%VENV_PYTHON%" -m trajectory_splitter.web_gui
if errorlevel 1 (
    echo.
    echo [trajectory_splitter web] launch failed, error code %errorlevel%
    echo check the venv path, or run: .venv\Scripts\pip install -r trajectory_splitter\requirements.txt
    pause
)
