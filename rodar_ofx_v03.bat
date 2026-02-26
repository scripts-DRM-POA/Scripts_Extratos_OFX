@echo off
REM Caminho do Python do Anaconda
SET PYTHON_PATH=C:\Anaconda\python.exe

REM Caminho do script processa_ofx_jupyter_v02.py
SET SCRIPT_PATH="K:\04_EPFI_Fiscalizacoes\Códigos_scrips\Processa_ofx\processa_ofx.py"

echo Rodando processamento OFX...

REM Executando o Python e chamando o script
"%PYTHON_PATH%" -c "import sys; sys.path.append(r'%SCRIPT_PATH%'); import processa_ofx_jupyter_v03; processa_ofx_jupyter_v03.process_dir()"

echo Processo finalizado.
pause