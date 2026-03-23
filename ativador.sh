#!/bin/bash
# Ativador - Gerenciamento de Telefones (Linux/Server)

cd "$(dirname "$0")"

echo ""
echo "============================================"
echo "  GERENCIAMENTO DE TELEFONES - Ativador"
echo "============================================"
echo ""

# Verificar Python
if ! command -v python3 &> /dev/null && ! command -v python &> /dev/null; then
    echo "[ERRO] Python nao encontrado!"
    echo "Instale: sudo apt install python3 python3-venv python3-pip  # Debian/Ubuntu"
    echo "Ou:      sudo yum install python3 python3-pip               # CentOS/RHEL"
    exit 1
fi

PYTHON_CMD=$(command -v python3 2>/dev/null || command -v python 2>/dev/null)
echo "[OK] Python encontrado: $($PYTHON_CMD --version)"
echo ""

# Criar ambiente virtual
if [ ! -f ".venv/bin/activate" ]; then
    echo "[1/3] Criando ambiente virtual..."
    $PYTHON_CMD -m venv .venv
    echo "[OK] Ambiente virtual criado."
else
    echo "[1/3] Ambiente virtual ja existe."
fi
echo ""

# Instalar dependências
echo "[2/3] Instalando dependencias..."
source .venv/bin/activate
pip install --upgrade pip -q
pip install -r requirements.txt -q
echo "[OK] Dependencias instaladas."
echo ""

# Iniciar aplicação
echo "[3/3] Iniciando aplicacao..."
echo ""
echo "Acesse: http://localhost:8501 (ou use o IP do servidor)"
echo "Pressione Ctrl+C para encerrar."
echo ""

exec streamlit run app.py --server.headless true --server.address 0.0.0.0
