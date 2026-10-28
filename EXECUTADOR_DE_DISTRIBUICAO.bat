@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Marine PGR - Gerar distribuicao
set "MARINE_PY="
py -3 --version >nul 2>&1 && set "MARINE_PY=py -3"
if not defined MARINE_PY (
  python --version >nul 2>&1 && set "MARINE_PY=python"
)
if not defined MARINE_PY (
  echo Python 3.10 ou superior nao encontrado. Instale-o para gerar a distribuicao.
  goto erro
)
if not exist ".build-env\Scripts\python.exe" (
  %MARINE_PY% -m venv .build-env
  if errorlevel 1 goto erro
)
".build-env\Scripts\python.exe" -m pip install -r requirements-build.txt
if errorlevel 1 goto erro
".build-env\Scripts\python.exe" scripts\gerar_distribuicao.py
if errorlevel 1 goto erro
echo.
echo PRONTO: o ZIP para os usuarios esta na pasta distribuicao.
echo Publique esse ZIP nos arquivos de uma Release do GitHub.
pause
exit /b 0
:erro
echo Nao foi possivel gerar a distribuicao. Confira a mensagem acima.
pause
exit /b 1
