@echo off
title Executor de Scripts - Auditoria
color 0A
setlocal EnableDelayedExpansion

pushd "%~dp0" 2>nul

set DIRETORIO_REDE=%CD%

if not exist "Logs" mkdir Logs

for /f "tokens=1-3 delims=/" %%a in ("%date%") do (
    set DIA=%%a
    set MES=%%b
    set ANO=%%c
)

for /f "tokens=1-2 delims=:" %%a in ("%time%") do (
    set HORA=%%a
    set MIN=%%b
)

set LOG=Logs\log_%USERNAME%_!ANO!!MES!!DIA!_!HORA!!MIN!.txt

echo ================================================== >> "!LOG!"
echo EXECUTOR DE SCRIPTS - AUDITORIA >> "!LOG!"
echo Usuario: %USERNAME% >> "!LOG!"
echo Maquina: %COMPUTERNAME% >> "!LOG!"
echo Data: %date% %time% >> "!LOG!"
echo ================================================== >> "!LOG!"
echo. >> "!LOG!"

set PYTHON_EXE=%USERPROFILE%\Appdata\Local\anaconda3\python.exe

if not exist "%PYTHON_EXE%" (
    echo ERRO: Anaconda nao encontrado. >> "!LOG!"
    echo Anaconda nao encontrado.
    type "!LOG!"
    pause
    exit /b
)

echo.
echo ==========================================
echo        EXECUTOR DE SCRIPTS
echo ==========================================
echo.

set COUNT=0

for %%f in (*.py) do (
    set /a COUNT+=1
    set SCRIPT!COUNT!=%%f
    echo !COUNT! - %%f
)

if !COUNT! EQU 0 (
    echo Nenhum script .py encontrado na pasta.
    pause
    exit /b
)

echo.
set /p ESCOLHA=Digite o numero do script desejado:

set SCRIPT_ESCOLHIDO=!SCRIPT%ESCOLHA%!

if "!SCRIPT_ESCOLHIDO!"=="" (
    echo Opcao invalida.
    pause
    exit /b
)

echo Script selecionado: !SCRIPT_ESCOLHIDO! >> "!LOG!"
echo ===== INICIO EXECUCAO ===== >> "!LOG!"
echo. >> "!LOG!"

"%PYTHON_EXE%" "!SCRIPT_ESCOLHIDO!" >> "!LOG!" 2>&1

echo. >> "!LOG!"
echo ===== FIM EXECUCAO ===== >> "!LOG!"

echo.
echo ==========================================
echo Execucao finalizada.
echo Log salvo em:
echo %DIRETORIO_REDE%\!LOG!
echo ==========================================
echo.

type "!LOG!"
pause

popd
endlocal