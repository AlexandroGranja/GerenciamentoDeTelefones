"""Inicializa o banco PostgreSQL do projeto."""

from urllib.parse import urlsplit

from src.core.config import get_database_url, is_postgres_configured
from src.db.repository import init_db


def _mask_database_url(url: str) -> str:
    try:
        parts = urlsplit(url)
        if parts.password:
            netloc = parts.netloc.replace(parts.password, "***")
            return parts._replace(netloc=netloc).geturl()
        return url
    except Exception:
        return url


def main() -> int:
    if not is_postgres_configured():
        print("DATABASE_URL nao configurado.")
        print("Crie um arquivo .env com base em .env.example ou defina a variavel manualmente.")
        return 1

    print("Inicializando schema PostgreSQL...")
    print(f"Conexao: {_mask_database_url(get_database_url())}")
    init_db()
    print("Schema PostgreSQL criado com sucesso.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
