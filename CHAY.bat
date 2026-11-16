@echo off
rem Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo Chua cai Python. Hay cai Python 3.11 tro len tai https://www.python.org/downloads/
  echo Nho tick "Add python.exe to PATH" khi cai.
  pause
  exit /b 1
)

if not exist .venv (
  echo Dang tao moi truong ao .venv ...
  python -m venv .venv
)
call .venv\Scripts\activate.bat
pip install -q -r requirements.txt

start "" cmd /c "timeout /t 3 >nul & start http://127.0.0.1:8000"
echo Dang chay tai http://127.0.0.1:8000  (Ctrl+C de dung)
python -m uvicorn app.main:app --host 127.0.0.1 --port 8000
pause
