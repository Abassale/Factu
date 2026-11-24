@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
set PY=python
where py >nul 2>nul && set PY=py
%PY% --version >nul 2>nul || (echo Python n'est pas installé : installe-le depuis https://www.python.org puis relance. & pause & exit /b 1)
echo Vérification des composants...
%PY% -m pip install --quiet --disable-pip-version-check -r outil\requirements.txt
if "%~1"=="" (
  %PY% outil\factu.py entree
) else (
  %PY% outil\factu.py entree --lot %1
)
if exist sortie start "" sortie
echo.
pause
