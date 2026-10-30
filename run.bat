@echo off
setlocal
cd /d "%~dp0"

rem 需要切换的 conda 环境名，按需改这里
set "CONDA_ENV=base"

rem ---- 激活运行环境：项目内 .venv -> venv -> conda ----
if exist ".venv\Scripts\activate.bat" goto use_venv
if exist "venv\Scripts\activate.bat" goto use_venv_dir

where conda >nul 2>nul
if errorlevel 1 goto no_env
call conda activate %CONDA_ENV% 2>nul
if errorlevel 1 goto no_env
echo [env] conda %CONDA_ENV% activated
goto launch

:use_venv
call ".venv\Scripts\activate.bat"
echo [env] .venv activated
goto launch

:use_venv_dir
call "venv\Scripts\activate.bat"
echo [env] venv activated
goto launch

:no_env
echo [env] no virtualenv found, using system Python

:launch
start "" cmd /c "timeout /t 2 >nul & start http://127.0.0.1:8000"
python -m uvicorn app:app --host 127.0.0.1 --port 8000

endlocal
