@echo off
chcp 65001 > nul
echo ============================================
echo   Airbnb 분석 CLI — PyInstaller 빌드
echo ============================================
echo.

set PYTHON=C:\Users\HOME\AppData\Local\Programs\Python\Python313\python.exe
set PYINST=%PYTHON% -m PyInstaller

echo [1/2] 이전 빌드 정리...
if exist dist\airbnb_CLI.exe del /q dist\airbnb_CLI.exe
if exist build\airbnb_CLI rmdir /s /q build\airbnb_CLI

echo [2/2] 빌드 시작...
%PYINST% ^
  --onefile ^
  --console ^
  --name "airbnb_CLI" ^
  --hidden-import xlsxwriter ^
  --collect-all curl_cffi ^
  cli_app.py

echo.
if exist dist\airbnb_CLI.exe (
  echo [OK] 빌드 성공: dist\airbnb_CLI.exe
) else (
  echo [ERROR] 빌드 실패 - 위 오류를 확인하세요.
)
pause
