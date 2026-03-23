@echo off
setlocal EnableDelayedExpansion

echo.
echo ============================================
echo   INICIADOR COMPLETO - Gerenciamento + Chamados
echo ============================================
echo.

set "ROOT=%~dp0"
set "BACKEND_DIR=%ROOT%\..\Sistema de Chamados TI\backend"
set "FRONTEND_DIR=%ROOT%\..\Sistema de Chamados TI\frontend"

REM =======================
REM Evita conflito de porta
REM =======================
REM Se já tiver algum serviço usando as portas, mata para recomeçar.
for /f "tokens=5" %%p in ('netstat -ano ^| findstr ":8000"') do taskkill /F /PID %%p >nul 2>&1
for /f "tokens=5" %%p in ('netstat -ano ^| findstr ":3000"') do taskkill /F /PID %%p >nul 2>&1
for /f "tokens=5" %%p in ('netstat -ano ^| findstr ":8501"') do taskkill /F /PID %%p >nul 2>&1

REM =======================
REM BACKEND Chamados (8000)
REM =======================
echo [1/3] Backend Chamados...
set "PY_BACKEND=%BACKEND_DIR%\venv\Scripts\python.exe"
if not exist "%PY_BACKEND%" (
  pushd "%BACKEND_DIR%"
  echo Criando venv backend...
  python -m venv venv
  venv\Scripts\pip.exe install --upgrade pip
  venv\Scripts\pip.exe install -r requirements.txt
  popd
)

pushd "%BACKEND_DIR%"
echo Iniciando API (8000)...
start "" /b venv\Scripts\python.exe -m uvicorn app.main:app --host 0.0.0.0 --port 8000
popd

timeout /t 3 /nobreak >nul

REM =======================
REM FRONTEND Chamados (3000)
REM =======================
echo [2/3] Frontend Chamados...
pushd "%FRONTEND_DIR%"
if not exist "node_modules" (
  echo Instalando dependencias frontend...
  npm install
)
echo Iniciando Web (3000)...
start "" /b node_modules\.bin\vite --host 0.0.0.0 --port 3000
popd

timeout /t 3 /nobreak >nul

REM =======================
REM GERENCIAMENTO (8501)
REM =======================
echo [3/3] Gerenciamento de Telefones...

set "PY_GER=%ROOT%\.venv\Scripts\python.exe"
if not exist "%PY_GER%" (
  pushd "%ROOT%"
  echo Criando ambiente virtual do Gerenciamento...
  python -m venv .venv
  .venv\Scripts\pip.exe install --upgrade pip -q
  .venv\Scripts\pip.exe install -r requirements.txt -q
  popd
)

pushd "%ROOT%"
echo Iniciando Streamlit (8501)...
start "" /b .venv\Scripts\python.exe -m streamlit run app.py --server.headless true --server.address 0.0.0.0 --server.port 8501
popd

echo.
echo Servicos iniciados.
echo.
echo - Gerenciamento: http://localhost:8501
echo - Chamados: http://localhost:3000
echo.
REM sem pause para funcionar em execucao nao-interativa

