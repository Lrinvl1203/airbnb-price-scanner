@echo off
chcp 65001 > nul
echo ============================================
echo   Booking.com 분석 CLI — PyInstaller 빌드
echo ============================================
echo.

set PYTHON=C:\Users\HOME\AppData\Local\Programs\Python\Python313\python.exe
set PYINST=%PYTHON% -m PyInstaller

echo [1/2] 이전 빌드 정리...
if exist dist\booking_CLI.exe del /q dist\booking_CLI.exe
if exist build\booking_CLI rmdir /s /q build\booking_CLI

echo [2/2] 빌드 시작...
%PYINST% ^
  --onefile ^
  --console ^
  --name "booking_CLI" ^
  --hidden-import xlsxwriter ^
  --hidden-import bs4 ^
  --collect-all curl_cffi ^
  --collect-all bs4 ^
  cli_app.py

echo.
if exist dist\booking_CLI.exe (
  echo [OK] 빌드 성공: dist\booking_CLI.exe
) else (
  echo [ERROR] 빌드 실패 - 위 오류를 확인하세요.
)
pause
