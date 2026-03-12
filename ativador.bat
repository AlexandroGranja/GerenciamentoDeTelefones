@echo off
setlocal EnableDelayedExpansion
title Ativador - Gerenciamento de Telefones

cd /d "%~dp0"

echo.
echo ============================================
echo   GERENCIAMENTO DE TELEFONES - Ativador
echo ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python nao encontrado!
    echo.
    echo Instale o Python em: https://www.python.org/downloads/
    echo Marque "Add Python to PATH" na instalacao.
    echo.
    pause
    exit /b 1
)

echo [OK] Python encontrado
python --version
echo.

if not exist ".venv\Scripts\python.exe" (
    echo [1/4] Criando ambiente virtual...
    python -m venv .venv
    if errorlevel 1 (
        echo [ERRO] Falha ao criar ambiente virtual.
        pause
        exit /b 1
    )
    echo [OK] Ambiente virtual criado.
) else (
    echo [1/4] Ambiente virtual ja existe.
)
echo.

echo [2/4] Instalando dependencias...
.venv\Scripts\python.exe -m pip install --upgrade pip -q
.venv\Scripts\pip.exe install -r requirements.txt -q
if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias.
    pause
    exit /b 1
)
echo [OK] Dependencias instaladas.
echo.

echo [3/4] Sincronizando banco de dados...
.venv\Scripts\python.exe -m scripts.sync_db 2>nul
if errorlevel 1 (
    echo [AVISO] Sincronizacao falhou. Sincronize em Config ao abrir o app.
) else (
    echo [OK] Banco sincronizado.
)
echo.

echo [4/4] Iniciando aplicacao...
echo.
echo Abrindo em: http://localhost:8501
echo Pressione Ctrl+C para encerrar.
echo.

.venv\Scripts\python.exe -m streamlit run app.py --server.headless true --server.address 0.0.0.0

pause
