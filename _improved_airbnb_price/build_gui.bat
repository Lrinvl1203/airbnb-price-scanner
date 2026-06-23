@echo off
chcp 65001 > nul
echo ============================================
echo   Airbnb 분석 GUI — PyInstaller 빌드
echo ============================================
echo.

set PYTHON=C:\Users\HOME\AppData\Local\Programs\Python\Python313\python.exe
set PYINST=%PYTHON% -m PyInstaller

echo [1/2] 이전 빌드 정리...
if exist dist\airbnb_GUI rmdir /s /q dist\airbnb_GUI
if exist build\airbnb_GUI rmdir /s /q build\airbnb_GUI

echo [2/2] 빌드 시작...
%PYINST% ^
  --onefile ^
  --windowed ^
  --name "airbnb_GUI" ^
  --hidden-import babel.numbers ^
  --hidden-import babel.dates ^
  --hidden-import tkcalendar ^
  --hidden-import ttkbootstrap ^
  --hidden-import xlsxwriter ^
  --collect-all ttkbootstrap ^
  --collect-all tkcalendar ^
  --collect-all curl_cffi ^
  --collect-data babel ^
  gui_app.py

echo.
if exist dist\airbnb_GUI.exe (
  echo [OK] 빌드 성공: dist\airbnb_GUI.exe
) else (
  echo [ERROR] 빌드 실패 - 위 오류를 확인하세요.
)
pause
