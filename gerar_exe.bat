@echo off
title Instalando dependencias e gerando .exe
color 0A
echo.
echo ================================================
echo  Verificador de CNPJ - Gerador de .exe
echo ================================================
echo.

:: Verifica Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python nao encontrado.
    echo Baixe em: https://www.python.org/downloads/
    echo Marque a opcao "Add Python to PATH" durante a instalacao.
    pause
    exit /b 1
)

echo [1/3] Instalando dependencias...
pip install pyinstaller requests openpyxl --quiet
if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias.
    pause
    exit /b 1
)
echo       OK

echo.
echo [2/3] Gerando o executavel (pode demorar 1-2 minutos)...
pyinstaller --onefile --windowed --name "VerificadorCNPJ" verificador_cnpj.py
if errorlevel 1 (
    echo [ERRO] Falha ao gerar o .exe.
    pause
    exit /b 1
)
echo       OK

echo.
echo [3/3] Movendo para a pasta atual...
if exist "dist\VerificadorCNPJ.exe" (
    copy /Y "dist\VerificadorCNPJ.exe" "VerificadorCNPJ.exe" >nul
    echo       OK
    echo.
    echo ================================================
    echo  PRONTO! Arquivo gerado: VerificadorCNPJ.exe
    echo ================================================
) else (
    echo [AVISO] Arquivo gerado em: dist\VerificadorCNPJ.exe
)

echo.
pause
