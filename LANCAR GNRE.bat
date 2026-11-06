@echo off
chcp 65001 >nul
setlocal
rem ==== endereco do robo no GitHub (usuario/repositorio) ====
set "REPO=NetoPianta/robo-gnre"
set "PASTA=%~dp0"
set "URL=https://raw.githubusercontent.com/%REPO%/main"

echo Verificando se ha versao nova do robo...
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { [Net.ServicePointManager]::SecurityProtocol='Tls12'; Invoke-WebRequest '%URL%/gnre_sped.py' -OutFile '%PASTA%gnre_sped.novo' -UseBasicParsing -TimeoutSec 30; Invoke-WebRequest '%URL%/requirements.txt' -OutFile '%PASTA%requirements.txt' -UseBasicParsing -TimeoutSec 30; exit 0 } catch { exit 1 }"
if errorlevel 1 (
  echo Nao consegui acessar o GitHub. Vou usar a versao que ja esta neste computador.
  if exist "%PASTA%gnre_sped.novo" del "%PASTA%gnre_sped.novo"
) else (
  move /y "%PASTA%gnre_sped.novo" "%PASTA%gnre_sped.py" >nul
  echo Robo atualizado.
)
if not exist "%PASTA%gnre_sped.py" (
  echo ERRO: o robo ainda nao foi baixado. Verifique a internet e tente de novo.
  pause
  exit /b
)

set "PY=python"
where python >nul 2>nul || set "PY=py"
%PY% --version >nul 2>nul
if errorlevel 1 (
  echo Python nao encontrado. Instale em https://www.python.org/downloads/ marcando "Add python.exe to PATH".
  pause
  exit /b
)

echo Conferindo os componentes (na primeira vez demora 1 a 2 minutos)...
%PY% -m pip install --quiet --disable-pip-version-check -r "%PASTA%requirements.txt"
%PY% "%PASTA%gnre_sped.py"
pause
