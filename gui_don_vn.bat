@echo off
chcp 65001 >nul
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
  echo Khong tim thay 'python'. Cai Python 3 roi chay lai file nay.
  pause
  exit /b 1
)
python gui_don_vn.py
pause