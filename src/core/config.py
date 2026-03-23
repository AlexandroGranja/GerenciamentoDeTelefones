"""Configurações e constantes da aplicação."""

import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

# Diretórios base
BASE_DIR = Path(__file__).resolve().parent.parent.parent
DATA_DIR = BASE_DIR / "data"
DOC_DIR = BASE_DIR / "doc"
DB_DIR = DATA_DIR / "db"

# Arquivos de configuração
RULES_FILE = "equipe_regras.csv"
EQUIPES_ALIMENTO_FILE = "equipes_alimento.csv"
EQUIPES_MEDICAMENTO_FILE = "equipes_medicamento.csv"

def get_db_path() -> Path:
    """Retorna caminho do banco SQLite legado/local."""
    DB_DIR.mkdir(parents=True, exist_ok=True)
    return DB_DIR / "gerenciamento_telefones.db"


def get_database_url() -> str:
    """Retorna URL do banco principal da aplicação."""
    return os.environ.get("DATABASE_URL", "").strip()


def get_database_backend() -> str:
    """Define o backend ativo: PostgreSQL quando DATABASE_URL existir."""
    return "postgres" if get_database_url() else "sqlite"


def is_postgres_configured() -> bool:
    """Indica se o projeto já está apontado para PostgreSQL."""
    return get_database_backend() == "postgres"


def get_chamados_app_url() -> str:
    """URL do sistema de chamados externo (separado)."""
    return os.environ.get("CHAMADOS_APP_URL", "").strip()
