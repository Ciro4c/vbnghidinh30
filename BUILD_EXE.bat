@echo off
rem Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
chcp 65001 >nul
cd /d "%~dp0"
title Dong goi VanBanND30.exe
rem Dung:  BUILD_EXE.bat          -> thu muc dist\VanBanND30\ (khuyen dung)
rem        BUILD_EXE.bat onefile  -> mot file dist\VanBanND30.exe

echo [1/5] Kiem tra Python...
where python >nul 2>nul
if errorlevel 1 (
  echo Chua cai Python. Cai Python 3.11 tro len tai https://www.python.org/downloads/
  pause & exit /b 1
)
python --version

echo.
echo [2/5] Tao moi truong ao .venv-build ^(lan dau mat ~30 giay^)...
if not exist .venv-build python -m venv .venv-build
if errorlevel 1 ( echo Tao moi truong ao that bai. & pause & exit /b 1 )
call .venv-build\Scripts\activate.bat

echo.
echo [3/5] Cai thu vien + PyInstaller ^(can mang, 1-5 phut, se hien thanh tien trinh^)...
pip install -r requirements-build.txt
if errorlevel 1 ( echo. & echo Cai thu vien that bai. Kiem tra ket noi mang roi chay lai. & pause & exit /b 1 )

set MODE=--onedir
if /i "%1"=="onefile" set MODE=--onefile

echo.
echo [4/5] Don dep ban build cu...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo [5/5] Dong goi ^(2-10 phut, cua so se in nhieu dong INFO, cu de yen^)...
pyinstaller --noconfirm --clean --noconsole %MODE% --name VanBanND30 ^
  --add-data "app\static;app\static" ^
  --add-data "app\profiles;app\profiles" ^
  --collect-submodules uvicorn ^
  --collect-data docx --collect-data pptx ^
  --version-file version_info.txt ^
  chay_app.py
if errorlevel 1 ( echo. & echo Dong goi that bai. Chup man hinh cua so nay gui lai. & pause & exit /b 1 )

echo.
echo ================= XONG =================
if /i "%1"=="onefile" ( echo File: %cd%\dist\VanBanND30.exe ) else ( echo Thu muc: %cd%\dist\VanBanND30\  ^(chay VanBanND30.exe trong do^) )
echo Copy sang may khac duoc, may do KHONG can cai Python.
pause
