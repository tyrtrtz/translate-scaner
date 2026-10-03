@echo off
setlocal
cd /d "%~dp0"
py -3 -m venv .venv
if errorlevel 1 goto fail
".venv\Scripts\python.exe" -m pip install -r requirements-build.txt
if errorlevel 1 goto fail
".venv\Scripts\python.exe" -m unittest -v
if errorlevel 1 goto fail
".venv\Scripts\python.exe" -m PyInstaller --clean --noconfirm --onefile --windowed --name HeaderChecker app.py
if errorlevel 1 goto fail
echo Built: dist\HeaderChecker.exe
pause
exit /b 0
:fail
echo Build failed. See the error above.
pause
exit /b 1
