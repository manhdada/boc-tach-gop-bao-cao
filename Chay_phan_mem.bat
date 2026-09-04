@echo off
setlocal
set "APP_DIR=%~dp0"
set "BUNDLED_PY=C:\Users\ADMIN\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%APP_DIR%BoctachvaGopBaoCao.exe" (
  start "" "%APP_DIR%BoctachvaGopBaoCao.exe"
  exit /b 0
)
if exist "%BUNDLED_PY%" (
  start "" "%BUNDLED_PY%" "%APP_DIR%run_app.py"
  exit /b 0
)
where pythonw >nul 2>nul
if %errorlevel%==0 (
  start "" pythonw "%APP_DIR%run_app.py"
  exit /b 0
)
echo Khong tim thay Python hoac file BoctachvaGopBaoCao.exe.
pause
