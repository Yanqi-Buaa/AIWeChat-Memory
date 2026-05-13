@echo off
setlocal
set "SKILL_DIR=%~dp0"
set "PYTHONPATH=%SKILL_DIR%"
C:\Users\Administrator\AppData\Local\Programs\Python\Python312\python.exe "%SKILL_DIR%scripts\run_ref_downloader.py" %*
