@echo off
chcp 65001 > nul
set "APP_DIR=%~dp0"
set "LOCAL_PY=C:\Users\HOME\AppData\Local\Programs\Python\Python313\python.exe"

if exist "%LOCAL_PY%" (
  "%LOCAL_PY%" "%APP_DIR%gui_app.py"
) else (
  py -3 "%APP_DIR%gui_app.py"
)
