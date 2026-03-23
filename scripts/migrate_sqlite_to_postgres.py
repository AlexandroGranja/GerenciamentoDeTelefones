"""Migra dados do SQLite legado para o PostgreSQL."""

import sqlite3
from pathlib import Path

import pandas as pd

from src.core.config import get_db_path, is_postgres_configured
from src.db.repository import LEGACY_COL_MAP, get_connection, init_db, save_linhas


def _sqlite_exists(path: Path) -> bool:
    return path.exists() and path.is_file()


def _fetch_table_df(conn: sqlite3.Connection, table_name: str) -> pd.DataFrame:
    try:
        return pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
    except Exception:
        return pd.DataFrame()


def _prepare_unique_linhas(df: pd.DataFrame, modo: str) -> tuple[pd.DataFrame, int, int]:
    """Mantem apenas o registro mais recente por linha dentro do modo."""
    if df.empty:
        return df.copy(), 0, 0

    prepared = df.copy()
    linha_col = "Linha" if "Linha" in prepared.columns else "linha"
    if "modo" not in prepared.columns:
        prepared["modo"] = ""
    if linha_col not in prepared.columns:
        prepared[linha_col] = ""

    prepared["modo"] = prepared["modo"].fillna("").astype(str).str.strip()
    prepared[linha_col] = prepared[linha_col].fillna("").astype(str).str.strip()
    prepared = prepared[prepared["modo"] == modo].copy()
    before_count = len(prepared)
    prepared = prepared[prepared[linha_col] != ""].copy()
    empty_removed = before_count - len(prepared)
    if prepared.empty:
        return prepared, before_count, empty_removed

    if "id" in prepared.columns:
        prepared["id"] = pd.to_numeric(prepared["id"], errors="coerce")
        prepared = prepared.sort_values([linha_col, "id"], na_position="first")
    else:
        prepared = prepared.sort_values([linha_col])

    prepared = prepared.drop_duplicates(subset=[linha_col], keep="last").copy()
    return prepared, before_count, empty_removed


def _migrate_users(sqlite_conn: sqlite3.Connection) -> int:
    users_df = _fetch_table_df(sqlite_conn, "usuarios")
    if users_df.empty:
        return 0

    migrated = 0
    pg_conn = get_connection()
    try:
        with pg_conn.cursor() as cur:
            for _, row in users_df.iterrows():
                username = str(row.get("username") or "").strip().lower()
                if not username:
                    continue
                cur.execute("SELECT 1 FROM usuarios_app WHERE username = %s", (username,))
                if cur.fetchone():
                    continue
                cur.execute(
                    """
                    INSERT INTO usuarios_app (username, password_hash, salt, auth_provider, is_admin, ativo)
                    VALUES (%s, %s, %s, %s, %s, %s)
                    """,
                    (
                        username,
                        str(row.get("password_hash") or ""),
                        str(row.get("salt") or ""),
                        "local",
                        bool(row.get("is_admin") or 0),
                        True,
                    ),
                )
                migrated += 1
        pg_conn.commit()
        return migrated
    finally:
        pg_conn.close()


def _migrate_linhas(sqlite_conn: sqlite3.Connection) -> tuple[int, int]:
    linhas_df = _fetch_table_df(sqlite_conn, "linhas")
    if linhas_df.empty:
        return 0, 0

    rev_map = {v: k for k, v in LEGACY_COL_MAP.items()}
    linhas_df = linhas_df.rename(columns={c: rev_map[c] for c in linhas_df.columns if c in rev_map})

    migrated_ativas = 0
    migrated_desativadas = 0
    if "modo" not in linhas_df.columns:
        return 0, 0

    df_ativas, ativas_before, ativas_empty_removed = _prepare_unique_linhas(linhas_df, "ativas")
    df_desativadas, desativadas_before, desativadas_empty_removed = _prepare_unique_linhas(linhas_df, "desativadas")

    ativas_dedup_removed = ativas_before - ativas_empty_removed - len(df_ativas)
    desativadas_dedup_removed = desativadas_before - desativadas_empty_removed - len(df_desativadas)

    print(
        "Linhas ativas: "
        f"{ativas_before} origem | {ativas_empty_removed} sem numero | "
        f"{ativas_dedup_removed} duplicadas removidas | {len(df_ativas)} unicas"
    )
    print(
        "Linhas desativadas: "
        f"{desativadas_before} origem | {desativadas_empty_removed} sem numero | "
        f"{desativadas_dedup_removed} duplicadas removidas | {len(df_desativadas)} unicas"
    )

    drop_cols = ["id", "criado_em", "modo"]
    if not df_ativas.empty:
        migrated_ativas = save_linhas(df_ativas.drop(columns=drop_cols, errors="ignore"), modo="ativas")
    if not df_desativadas.empty:
        migrated_desativadas = save_linhas(df_desativadas.drop(columns=drop_cols, errors="ignore"), modo="desativadas")
    return migrated_ativas, migrated_desativadas


def main() -> int:
    if not is_postgres_configured():
        print("DATABASE_URL nao configurado. Configure o PostgreSQL antes de migrar.")
        return 1

    sqlite_path = get_db_path()
    if not _sqlite_exists(sqlite_path):
        print(f"Banco SQLite legado nao encontrado: {sqlite_path}")
        return 1

    print("Inicializando schema PostgreSQL...")
    init_db()

    sqlite_conn = sqlite3.connect(str(sqlite_path))
    try:
        print(f"Migrando a partir de: {sqlite_path}")
        migrated_users = _migrate_users(sqlite_conn)
        migrated_ativas, migrated_desativadas = _migrate_linhas(sqlite_conn)
    finally:
        sqlite_conn.close()

    print(f"Usuarios migrados: {migrated_users}")
    print(f"Linhas ativas migradas: {migrated_ativas}")
    print(f"Linhas desativadas migradas: {migrated_desativadas}")
    print("Migracao concluida.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
