@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>&1
if %errorlevel% equ 0 (
  set "PY=py -3"
) else (
  where python >nul 2>&1
  if %errorlevel% equ 0 (set "PY=python") else (
    echo Instale Python 3.9+ de https://www.python.org/downloads/windows/
    echo Marque "tcl/tk and IDLE" no instalador.
    exit /b 1
  )
)

echo Instalando PyInstaller...
%PY% -m pip install --upgrade pip
%PY% -m pip install -r requirements-build.txt
echo Gerando executavel...
%PY% -m PyInstaller BuscaBranchGit_v3.spec
if %errorlevel% neq 0 (
  echo Falha no PyInstaller.
  exit /b 1
)
echo.
echo Pronto: dist\BuscaBranchGit_v3.exe
start "" "dist"
endlocal
