@echo off
cd /d "%~dp0"
python -m pip install --upgrade pip pyinstaller -q
python -m PyInstaller --onefile --windowed --name BuscaBranchGit --clean --noconfirm ^
  --hidden-import=git ^
  --hidden-import=gitdb ^
  --hidden-import=gitdb.db ^
  --hidden-import=git.cmd ^
  index.py
if errorlevel 1 exit /b 1
echo.
echo Executavel: %~dp0dist\BuscaBranchGit.exe
pause
