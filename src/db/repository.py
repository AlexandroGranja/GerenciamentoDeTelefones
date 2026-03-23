"""Repositório para acesso ao banco de dados."""

import hashlib
import json
import secrets
import sqlite3
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Optional

import pandas as pd

from src.core.config import get_database_backend, get_database_url, get_db_path

try:
    import psycopg
    from psycopg.errors import UniqueViolation
    from psycopg.errors import DeadlockDetected
except ImportError:  # pragma: no cover - ambiente sem psycopg
    psycopg = None
    UniqueViolation = Exception
    DeadlockDetected = Exception


LEGACY_COL_MAP = {
    "↕ Ordem": "ordem_manual",
    "Codigo": "codigo",
    "Nome": "nome",
    "Equipe": "equipe",
    "EquipePadrao": "equipe_padrao",
    "GrupoEquipe": "grupo_equipe",
    "TipoEquipe": "tipo_equipe",
    "Localidade": "localidade",
    "Data da Troca": "data_troca",
    "Data Retorno": "data_retorno",
    "Data Ocorrência": "data_ocorrencia",
    "Data Solicitação TBS": "data_solicitacao_tbs",
    "Gestor": "gestor",
    "Supervisor": "supervisor",
    "Segmento": "segmento",
    "Papel": "papel",
    "Linha": "linha",
    "E-mail": "email",
    "Gerenciamento": "gerenciamento",
    "IMEI A": "imei_a",
    "IMEI B": "imei_b",
    "CHIP": "chip",
    "Marca": "marca",
    "Aparelho": "aparelho",
    "Modelo": "modelo",
    "Setor": "setor",
    "Cargo": "cargo",
    "Desconto": "desconto",
    "Perfil": "perfil",
    "Empresa": "empresa",
    "Ativo": "ativo",
    "Numero de Serie": "numero_serie",
    "Patrimonio": "patrimonio",
    "Operadora": "operadora",
    "Nome de Guerra": "nome_guerra",
    "Motivo": "motivo",
    "Observação": "observacao",
    "Aba": "aba",
}

REV_COL_MAP = {v: k for k, v in LEGACY_COL_MAP.items()}


def _hash_password(password: str, salt: str) -> str:
    """Gera hash da senha com salt."""
    return hashlib.sha256((salt + password).encode("utf-8")).hexdigest()


def _backend() -> str:
    return get_database_backend()


def _is_postgres() -> bool:
    return _backend() == "postgres"


def _normalize_record_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    text = str(value).strip() if not isinstance(value, (dict, list, bool, int)) else value
    if isinstance(text, str) and not text:
        return None
    return text


def _connect_postgres():
    if psycopg is None:
        raise RuntimeError("psycopg nao esta instalado para uso com PostgreSQL.")
    db_url = get_database_url()
    if not db_url:
        raise RuntimeError("DATABASE_URL nao configurado.")
    return psycopg.connect(db_url)


def get_connection(db_path: Optional[Path] = None):
    """Retorna conexão do backend ativo."""
    if _is_postgres():
        return _connect_postgres()
    path = db_path or get_db_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path))
    conn.row_factory = sqlite3.Row
    return conn


def criar_usuario(username: str, password: str, is_admin: bool = False, db_path: Optional[Path] = None) -> bool:
    """Cria novo usuário. Retorna True se sucesso."""
    salt = secrets.token_hex(16)
    password_hash = _hash_password(password, salt)
    conn = get_connection(db_path)
    try:
        if _is_postgres():
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO usuarios_app (username, password_hash, salt, auth_provider, is_admin)
                    VALUES (%s, %s, %s, %s, %s)
                    """,
                    (username.strip().lower(), password_hash, salt, "local", bool(is_admin)),
                )
            conn.commit()
            return True
        conn.execute(
            "INSERT INTO usuarios (username, password_hash, salt, is_admin) VALUES (?, ?, ?, ?)",
            (username.strip().lower(), password_hash, salt, 1 if is_admin else 0),
        )
        conn.commit()
        return True
    except (sqlite3.IntegrityError, UniqueViolation):
        return False
    finally:
        conn.close()


def verificar_login(username: str, password: str, db_path: Optional[Path] = None) -> Optional[dict]:
    """Verifica credenciais. Retorna dict com username, is_admin ou None."""
    conn = get_connection(db_path)
    try:
        if _is_postgres():
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT username, password_hash, salt, is_admin
                    FROM usuarios_app
                    WHERE username = %s AND ativo = TRUE
                    """,
                    (username.strip().lower(),),
                )
                row = cur.fetchone()
        else:
            row = conn.execute(
                "SELECT username, password_hash, salt, is_admin FROM usuarios WHERE username = ?",
                (username.strip().lower(),),
            ).fetchone()
        if not row:
            return None
        _, stored_hash, salt, is_admin = row
        if _hash_password(password, salt) == stored_hash:
            return {"username": row[0], "is_admin": bool(is_admin)}
        return None
    finally:
        conn.close()


def listar_usuarios(db_path: Optional[Path] = None) -> list[dict]:
    """Lista todos os usuários (sem senha)."""
    conn = get_connection(db_path)
    try:
        if _is_postgres():
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT id, username, is_admin, criado_em FROM usuarios_app ORDER BY username"
                )
                rows = cur.fetchall()
        else:
            rows = conn.execute(
                "SELECT id, username, is_admin, criado_em FROM usuarios ORDER BY username"
            ).fetchall()
        return [
            {"id": r[0], "username": r[1], "is_admin": bool(r[2]), "criado_em": r[3]}
            for r in rows
        ]
    finally:
        conn.close()


def excluir_usuario(username: str, db_path: Optional[Path] = None) -> bool:
    """Exclui usuário. Retorna True se excluiu."""
    conn = get_connection(db_path)
    try:
        if _is_postgres():
            with conn.cursor() as cur:
                cur.execute("DELETE FROM usuarios_app WHERE username = %s", (username.strip().lower(),))
                deleted = cur.rowcount > 0
        else:
            cur = conn.execute("DELETE FROM usuarios WHERE username = ?", (username.strip().lower(),))
            deleted = cur.rowcount > 0
        conn.commit()
        return deleted
    finally:
        conn.close()


def atualizar_senha_usuario(username: str, nova_senha: str, db_path: Optional[Path] = None) -> bool:
    """Atualiza a senha de um usuário. Retorna True se atualizou."""
    user = username.strip().lower()
    if not user or len(nova_senha or "") < 4:
        return False
    salt = secrets.token_hex(16)
    password_hash = _hash_password(nova_senha, salt)
    conn = get_connection(db_path)
    try:
        if _is_postgres():
            with conn.cursor() as cur:
                cur.execute(
                    """
                    UPDATE usuarios_app
                    SET password_hash = %s, salt = %s, ativo = TRUE
                    WHERE username = %s
                    """,
                    (password_hash, salt, user),
                )
                updated = cur.rowcount > 0
        else:
            cur = conn.execute(
                "UPDATE usuarios SET password_hash = ?, salt = ? WHERE username = ?",
                (password_hash, salt, user),
            )
            updated = cur.rowcount > 0
        conn.commit()
        return updated
    finally:
        conn.close()


def tem_usuarios(db_path: Optional[Path] = None) -> bool:
    """Verifica se existe pelo menos um usuário."""
    conn = get_connection(db_path)
    try:
        if _is_postgres():
            with conn.cursor() as cur:
                cur.execute("SELECT COUNT(*) FROM usuarios_app")
                return cur.fetchone()[0] > 0
        return conn.execute("SELECT COUNT(*) FROM usuarios").fetchone()[0] > 0
    finally:
        conn.close()


def obter_usuario_por_username(username: str, db_path: Optional[Path] = None) -> Optional[dict]:
    """Retorna usuário por username (sem senha). Para restaurar sessão."""
    conn = get_connection(db_path)
    try:
        if _is_postgres():
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT username, is_admin FROM usuarios_app WHERE username = %s",
                    (username.strip().lower(),),
                )
                row = cur.fetchone()
        else:
            row = conn.execute(
                "SELECT username, is_admin FROM usuarios WHERE username = ?",
                (username.strip().lower(),),
            ).fetchone()
        if not row:
            return None
        return {"username": row[0], "is_admin": bool(row[1])}
    finally:
        conn.close()


def obter_usuario_app_id(username: str, db_path: Optional[Path] = None) -> Optional[int]:
    """Retorna o id do usuário em `usuarios_app` (PostgreSQL) / `usuarios` (SQLite)."""
    username_norm = (username or "").strip().lower()
    if not username_norm:
        return None

    conn = get_connection(db_path)
    try:
        if _is_postgres():
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT id FROM usuarios_app WHERE username = %s LIMIT 1",
                    (username_norm,),
                )
                row = cur.fetchone()
        else:
            row = conn.execute(
                "SELECT id FROM usuarios WHERE username = ? LIMIT 1",
                (username_norm,),
            ).fetchone()

        if not row:
            return None
        return int(row[0])
    finally:
        conn.close()


def criar_sso_code(
    usuario_app_id: int,
    expira_em: datetime,
    db_path: Optional[Path] = None,
) -> Optional[str]:
    """
    Cria um código de SSO de 1 uso para redirecionar do Gerenciamento (Streamlit)
    para o sistema React externo.
    """
    if not _is_postgres():
        # Neste projeto o SSO unificado foi desenhado para PostgreSQL.
        return None

    if usuario_app_id is None:
        return None

    usuario_app_id_int = int(usuario_app_id)
    if not isinstance(expira_em, datetime):
        return None

    # token aleatorio curto mas suficiente; evitamos caracteres que atrapalham URL.
    code = secrets.token_hex(32)

    conn = get_connection(db_path)
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO sso_codes (code, usuario_id, expira_em, usado_em)
                VALUES (%s, %s, %s, NULL)
                """,
                (code, usuario_app_id_int, expira_em),
            )
        conn.commit()
        return code
    except UniqueViolation:
        # Caso o token colida (muito improvável), tenta gerar outro.
        return criar_sso_code(usuario_app_id_int, expira_em, db_path=db_path)
    finally:
        conn.close()


def criar_sessao(token: str, username: str, dias_validade: int = 30, db_path: Optional[Path] = None) -> None:
    """Cria sessão para login persistente (login único por usuário)."""
    expira_dt = datetime.now() + timedelta(days=dias_validade)
    expira_sqlite = expira_dt.strftime("%Y-%m-%d %H:%M:%S")
    conn = get_connection(db_path)
    try:
        if _is_postgres():
            with conn.cursor() as cur:
                cur.execute("SELECT id FROM usuarios_app WHERE username = %s", (username.strip().lower(),))
                row = cur.fetchone()
                if not row:
                    return
                usuario_id = row[0]
                cur.execute("DELETE FROM sessoes WHERE usuario_id = %s", (usuario_id,))
                cur.execute(
                    """
                    INSERT INTO sessoes (token, usuario_id, expira_em)
                    VALUES (%s, %s, %s)
                    ON CONFLICT (token) DO UPDATE SET usuario_id = EXCLUDED.usuario_id, expira_em = EXCLUDED.expira_em
                    """,
                    (token, usuario_id, expira_dt),
                )
        else:
            conn.execute("DELETE FROM sessoes WHERE username = ?", (username.strip().lower(),))
            conn.execute(
                "INSERT OR REPLACE INTO sessoes (token, username, expira_em) VALUES (?, ?, ?)",
                (token, username.strip().lower(), expira_sqlite),
            )
        conn.commit()
    finally:
        conn.close()


def validar_sessao(token: str, db_path: Optional[Path] = None) -> Optional[dict]:
    """Valida token de sessão. Retorna dados do usuário ou None."""
    conn = get_connection(db_path)
    try:
        if _is_postgres():
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT u.username, u.is_admin
                    FROM sessoes s
                    JOIN usuarios_app u ON u.id = s.usuario_id
                    WHERE s.token = %s AND s.expira_em > NOW()
                    """,
                    (token,),
                )
                row = cur.fetchone()
        else:
            row = conn.execute(
                "SELECT s.username, u.is_admin FROM sessoes s "
                "JOIN usuarios u ON u.username = s.username "
                "WHERE s.token = ? AND datetime(s.expira_em) > datetime('now')",
                (token,),
            ).fetchone()
        if not row:
            return None
        return {"username": row[0], "is_admin": bool(row[1])}
    finally:
        conn.close()


def encerrar_sessao(token: str, db_path: Optional[Path] = None) -> None:
    """Remove sessão (logout)."""
    conn = get_connection(db_path)
    try:
        if _is_postgres():
            with conn.cursor() as cur:
                cur.execute("DELETE FROM sessoes WHERE token = %s", (token,))
        else:
            conn.execute("DELETE FROM sessoes WHERE token = ?", (token,))
        conn.commit()
    finally:
        conn.close()


def init_db(db_path: Optional[Path] = None) -> None:
    """Inicializa o schema do banco."""
    schema_name = "schema_postgres.sql" if _is_postgres() else "schema.sql"
    schema_path = Path(__file__).parent / schema_name
    conn = get_connection(db_path)
    try:
        if _is_postgres():
            # Evita deadlocks quando duas instâncias do Streamlit tentam inicializar
            # o schema/alterações ao mesmo tempo.
            lock_key = "gerenciamento_telefones_init_db_v1"
            max_attempts = 3
            for attempt in range(1, max_attempts + 1):
                try:
                    with conn.cursor() as cur:
                        cur.execute("SELECT pg_advisory_lock(hashtext(%s))", (lock_key,))
                        try:
                            cur.execute(schema_path.read_text(encoding="utf-8"))
                            cur.execute("ALTER TABLE linhas ADD COLUMN IF NOT EXISTS ordem_manual INTEGER")
                            # Compatibilidade com a UI do sistema de chamados (React)
                            cur.execute("ALTER TABLE chamados ADD COLUMN IF NOT EXISTS category VARCHAR(100)")
                            cur.execute("ALTER TABLE chamados ADD COLUMN IF NOT EXISTS subcategory VARCHAR(100)")
                            cur.execute("ALTER TABLE chamados ADD COLUMN IF NOT EXISTS location VARCHAR(100)")
                            cur.execute("ALTER TABLE chamados ADD COLUMN IF NOT EXISTS equipment_info TEXT")
                            cur.execute("ALTER TABLE chamados ADD COLUMN IF NOT EXISTS internal_notes TEXT")
                            conn.commit()
                            return
                        finally:
                            cur.execute("SELECT pg_advisory_unlock(hashtext(%s))", (lock_key,))
                except DeadlockDetected:
                    # Deadlock em migrações raramente acontece mesmo com advisory lock
                    # (ex.: outras migrações externas). Faz retry com backoff simples.
                    if attempt >= max_attempts:
                        raise
                    conn.rollback()
                    sleep_s = 0.6 * attempt
                    time.sleep(sleep_s)

        conn.executescript(schema_path.read_text(encoding="utf-8"))
        cols = {r[1] for r in conn.execute("PRAGMA table_info(linhas)").fetchall()}
        if "data_troca" not in cols:
            conn.execute("ALTER TABLE linhas ADD COLUMN data_troca TEXT")
        if "data_retorno" not in cols:
            conn.execute("ALTER TABLE linhas ADD COLUMN data_retorno TEXT")
        if "data_ocorrencia" not in cols:
            conn.execute("ALTER TABLE linhas ADD COLUMN data_ocorrencia TEXT")
        if "data_solicitacao_tbs" not in cols:
            conn.execute("ALTER TABLE linhas ADD COLUMN data_solicitacao_tbs TEXT")
        if "marca" not in cols:
            conn.execute("ALTER TABLE linhas ADD COLUMN marca TEXT")
        if "patrimonio" not in cols:
            conn.execute("ALTER TABLE linhas ADD COLUMN patrimonio TEXT")
        if "nome_guerra" not in cols:
            conn.execute("ALTER TABLE linhas ADD COLUMN nome_guerra TEXT")
        if "motivo" not in cols:
            conn.execute("ALTER TABLE linhas ADD COLUMN motivo TEXT")
        if "observacao" not in cols:
            conn.execute("ALTER TABLE linhas ADD COLUMN observacao TEXT")
        audit_cols = {r[1] for r in conn.execute("PRAGMA table_info(audit_log)").fetchall()}
        if "chamado_id" not in audit_cols:
            conn.execute("ALTER TABLE audit_log ADD COLUMN chamado_id TEXT")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_audit_chamado_id ON audit_log(chamado_id)")
        conn.commit()
    finally:
        conn.close()


def save_linhas(df: pd.DataFrame, modo: str = "ativas", db_path: Optional[Path] = None) -> int:
    """Salva DataFrame na tabela `linhas`.

    No PostgreSQL, não removemos tudo da tabela para evitar violação de FK com
    `chamados.linha_id` (que referencia `linhas.id`). Em vez disso, fazemos UPSERT
    por `numero_linha` e, quando necessário, deletamos apenas linhas que não
    vieram no DataFrame e que não estejam referenciadas por chamados.
    """
    conn = get_connection(db_path)
    try:
        available = [c for c in LEGACY_COL_MAP if c in df.columns]
        df_export = df[available].copy() if available else pd.DataFrame()
        df_export = df_export.rename(columns=LEGACY_COL_MAP)
        if _is_postgres():
            with conn.cursor() as cur:
                records: list[dict[str, Any]] = []
                if not df_export.empty:
                    for record in df_export.to_dict("records"):
                        rec = {k: _normalize_record_value(v) for k, v in record.items()}
                        # `ordem_manual` precisa ser inteiro no PostgreSQL.
                        if rec.get("ordem_manual") not in (None, ""):
                            try:
                                rec["ordem_manual"] = int(float(str(rec.get("ordem_manual")).strip()))
                            except Exception:
                                rec["ordem_manual"] = None
                        rec["modo"] = modo
                        rec["modo_operacao"] = modo
                        rec["numero_linha"] = rec.get("linha")
                        rec["codigo_usuario_snapshot"] = rec.get("codigo")
                        rec["nome_usuario_snapshot"] = rec.get("nome")
                        rec["email_snapshot"] = rec.get("email")
                        # `numero_linha` é UNIQUE/NOT NULL: se vier vazio, não tenta persistir.
                        if rec.get("numero_linha") in (None, ""):
                            continue
                        records.append(rec)

                if records:
                    insert_cols = sorted({k for rec in records for k in rec.keys()})
                    placeholders = ", ".join(["%s"] * len(insert_cols))
                    columns_sql = ", ".join(insert_cols)

                    # UPSERT por `numero_linha` (evita trocar o `id`, mantendo FK em `chamados`).
                    update_cols = [c for c in insert_cols if c != "numero_linha"]
                    update_set_sql = ", ".join([f"{c} = EXCLUDED.{c}" for c in update_cols])
                    values = [tuple(rec.get(col) for col in insert_cols) for rec in records]

                    cur.executemany(
                        f"""
                        INSERT INTO linhas ({columns_sql})
                        VALUES ({placeholders})
                        ON CONFLICT (numero_linha) DO UPDATE SET {update_set_sql}
                        """.strip(),
                        values,
                    )

                # Deleta apenas o que não veio no DataFrame E não está referenciado por `chamados`.
                # Isso substitui o DELETE "total" anterior, evitando ForeignKeyViolation.
                keys = [r.get("numero_linha") for r in records if r.get("numero_linha") not in (None, "")]
                if keys:
                    placeholders_keys = ", ".join(["%s"] * len(keys))
                    cur.execute(
                        f"""
                        DELETE FROM linhas l
                        WHERE COALESCE(l.modo, l.modo_operacao) = %s
                          AND l.numero_linha NOT IN ({placeholders_keys})
                          AND NOT EXISTS (
                            SELECT 1 FROM chamados c
                            WHERE c.linha_id = l.id
                          )
                        """.strip(),
                        [modo] + keys,
                    )
                else:
                    # Se não há registros para persistir, não deletamos nada que esteja referenciado.
                    cur.execute(
                        """
                        DELETE FROM linhas l
                        WHERE COALESCE(l.modo, l.modo_operacao) = %s
                          AND NOT EXISTS (
                            SELECT 1 FROM chamados c
                            WHERE c.linha_id = l.id
                          )
                        """.strip(),
                        (modo,),
                    )
            conn.commit()
            return len(df_export)

        conn.execute("DELETE FROM linhas WHERE modo = ?", (modo,))
        if not df_export.empty:
            df_export["modo"] = modo
            df_export.to_sql("linhas", conn, if_exists="append", index=False)
        conn.commit()
        return len(df_export)
    finally:
        conn.close()


def load_linhas(modo: str = "ativas", db_path: Optional[Path] = None) -> pd.DataFrame:
    """Carrega linhas do banco. Retorna DataFrame vazio se não houver dados."""
    conn = get_connection(db_path)
    try:
        if _is_postgres():
            cols_sql = ", ".join(sorted(set(LEGACY_COL_MAP.values())))
            with conn.cursor() as cur:
                cur.execute(
                    f"SELECT {cols_sql} FROM linhas WHERE COALESCE(modo, modo_operacao) = %s",
                    (modo,),
                )
                rows = cur.fetchall()
                if not rows:
                    return pd.DataFrame()
                df = pd.DataFrame(rows, columns=sorted(set(LEGACY_COL_MAP.values())))
        else:
            df = pd.read_sql_query(
                "SELECT * FROM linhas WHERE modo = ?",
                conn,
                params=(modo,),
            )
            if df.empty:
                return pd.DataFrame()
            df = df.drop(columns=[c for c in df.columns if c in ("id", "criado_em", "modo")], errors="ignore")
        df = df.rename(columns={c: REV_COL_MAP[c] for c in df.columns if c in REV_COL_MAP})
        return df
    finally:
        conn.close()


def save_relacao_ativas(linhas: set[str], db_path: Optional[Path] = None) -> int:
    """Salva conjunto de linhas ativas na relação."""
    if _is_postgres():
        return len(linhas)
    conn = get_connection(db_path)
    try:
        conn.execute("DELETE FROM relacao_ativas")
        for ln in linhas:
            conn.execute("INSERT OR REPLACE INTO relacao_ativas (linha) VALUES (?)", (ln,))
        conn.commit()
        return len(linhas)
    finally:
        conn.close()


def load_relacao_ativas(db_path: Optional[Path] = None) -> frozenset[str]:
    """Carrega conjunto de linhas ativas da relação."""
    if _is_postgres():
        return frozenset()
    conn = get_connection(db_path)
    try:
        rows = conn.execute("SELECT linha FROM relacao_ativas").fetchall()
        return frozenset(r[0] for r in rows if r[0])
    finally:
        conn.close()


def has_data(modo: str = "ativas", db_path: Optional[Path] = None) -> bool:
    """Verifica se existem dados no banco para o modo."""
    conn = get_connection(db_path)
    try:
        if _is_postgres():
            with conn.cursor() as cur:
                cur.execute("SELECT COUNT(*) FROM linhas WHERE COALESCE(modo, modo_operacao) = %s", (modo,))
                return cur.fetchone()[0] > 0
        count = conn.execute("SELECT COUNT(*) FROM linhas WHERE modo = ?", (modo,)).fetchone()[0]
        return count > 0
    finally:
        conn.close()


def registrar_auditoria(
    acao: str,
    entidade: str,
    chave_registro: str = "",
    chamado_id: str = "",
    antes: Optional[dict] = None,
    depois: Optional[dict] = None,
    detalhes: str = "",
    user_id: str = "",
    username: str = "",
    origem: str = "app",
    db_path: Optional[Path] = None,
) -> None:
    """Registra um evento de auditoria."""
    conn = get_connection(db_path)
    try:
        before_json = json.dumps(antes, ensure_ascii=False) if antes is not None else None
        after_json = json.dumps(depois, ensure_ascii=False) if depois is not None else None
        if _is_postgres():
            chamado_val = int(chamado_id) if str(chamado_id or "").strip().isdigit() else None
            user_val = int(user_id) if str(user_id or "").strip().isdigit() else None
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO auditoria
                    (acao, entidade, entidade_id, chamado_id, antes_json, depois_json, detalhes, user_id, username, origem)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """,
                    (
                        acao,
                        entidade,
                        (chave_registro or "").strip(),
                        chamado_val,
                        before_json,
                        after_json,
                        (detalhes or "").strip(),
                        user_val,
                        (username or "").strip(),
                        (origem or "app").strip(),
                    ),
                )
        else:
            conn.execute(
                """
                INSERT INTO audit_log
                (acao, entidade, chave_registro, chamado_id, antes_json, depois_json, detalhes, user_id, username, origem)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    acao,
                    entidade,
                    (chave_registro or "").strip(),
                    (chamado_id or "").strip(),
                    before_json,
                    after_json,
                    (detalhes or "").strip(),
                    (user_id or "").strip(),
                    (username or "").strip(),
                    (origem or "app").strip(),
                ),
            )
        conn.commit()
    finally:
        conn.close()


def listar_auditoria(
    limit: int = 200,
    username: str = "",
    acao: str = "",
    entidade: str = "",
    chamado_id: str = "",
    db_path: Optional[Path] = None,
) -> list[dict]:
    """Lista eventos de auditoria mais recentes (com filtros opcionais)."""
    conn = get_connection(db_path)
    try:
        clauses: list[str] = []
        params: list[object] = []
        placeholder = "%s" if _is_postgres() else "?"
        if username.strip():
            clauses.append(f"username = {placeholder}")
            params.append(username.strip())
        if acao.strip():
            clauses.append(f"acao = {placeholder}")
            params.append(acao.strip())
        if entidade.strip():
            clauses.append(f"entidade = {placeholder}")
            params.append(entidade.strip())
        if chamado_id.strip():
            if _is_postgres() and chamado_id.strip().isdigit():
                clauses.append(f"chamado_id = {placeholder}")
                params.append(int(chamado_id.strip()))
            elif not _is_postgres():
                clauses.append(f"chamado_id = {placeholder}")
                params.append(chamado_id.strip())
        where_sql = f"WHERE {' AND '.join(clauses)}" if clauses else ""
        if _is_postgres():
            sql = f"""
                SELECT id, acao, entidade, entidade_id AS chave_registro, chamado_id, antes_json, depois_json, detalhes, user_id, username, origem, criado_em
                FROM auditoria
                {where_sql}
                ORDER BY id DESC
                LIMIT {placeholder}
            """
            params.append(max(1, int(limit)))
            with conn.cursor() as cur:
                cur.execute(sql, params)
                rows = cur.fetchall()
        else:
            sql = f"""
                SELECT id, acao, entidade, chave_registro, chamado_id, antes_json, depois_json, detalhes, user_id, username, origem, criado_em
                FROM audit_log
                {where_sql}
                ORDER BY id DESC
                LIMIT {placeholder}
            """
            params.append(max(1, int(limit)))
            rows = conn.execute(sql, params).fetchall()
        return [
            {
                "id": r[0],
                "acao": r[1],
                "entidade": r[2],
                "chave_registro": r[3],
                "chamado_id": r[4],
                "antes_json": r[5],
                "depois_json": r[6],
                "detalhes": r[7],
                "user_id": r[8],
                "username": r[9],
                "origem": r[10],
                "criado_em": r[11],
            }
            for r in rows
        ]
    finally:
        conn.close()


def garantir_chamado_stub(
    chamado_id: str,
    tipo: str = "gerenciamento",
    db_path: Optional[Path] = None,
) -> Optional[int]:
    """
    Garante que exista um registro mínimo na tabela `chamados` (PostgreSQL).

    Objetivo: permitir inserções futuras em tabelas dependentes (ex.: `chamado_eventos`)
    sem violar FKs quando o app recebe `chamado_id` por URL.
    """
    if not _is_postgres():
        return None
    chamado_id = str(chamado_id or "").strip()
    if not chamado_id:
        return None
    chamado_int = int(chamado_id) if chamado_id.isdigit() else None

    conn = get_connection(db_path)
    try:
        with conn.cursor() as cur:
            if chamado_int is not None:
                # Auditoria no app referencia `chamados(id)` usando `chamado_id` como int.
                cur.execute("SELECT id FROM chamados WHERE id = %s", (chamado_int,))
                row = cur.fetchone()
                if row:
                    return int(row[0])

                cur.execute(
                    """
                    INSERT INTO chamados (id, numero_chamado, tipo)
                    VALUES (%s, %s, %s)
                    RETURNING id
                    """,
                    (chamado_int, chamado_id, tipo),
                )
                new_id = cur.fetchone()
                conn.commit()
                return int(new_id[0]) if new_id else None
            else:
                cur.execute("SELECT id FROM chamados WHERE numero_chamado = %s", (chamado_id,))
                row = cur.fetchone()
                if row:
                    return int(row[0])

                cur.execute(
                    """
                    INSERT INTO chamados (numero_chamado, tipo)
                    VALUES (%s, %s)
                    RETURNING id
                    """,
                    (chamado_id, tipo),
                )
                new_id = cur.fetchone()
                conn.commit()
                return int(new_id[0]) if new_id else None
    except UniqueViolation:  # pragma: no cover (caso raro, concorrência)
        with conn.cursor() as cur:
            if chamado_int is not None:
                cur.execute("SELECT id FROM chamados WHERE id = %s", (chamado_int,))
            else:
                cur.execute("SELECT id FROM chamados WHERE numero_chamado = %s", (chamado_id,))
            row = cur.fetchone()
            return int(row[0]) if row else None
    finally:
        conn.close()


def vincular_chamado_linha(
    chamado_id: str,
    linha_num: str,
    db_path: Optional[Path] = None,
) -> None:
    """Vincula `chamados.linha_id` a partir de `linhas.numero_linha` (PostgreSQL)."""
    if not _is_postgres():
        return

    chamado_id = str(chamado_id or "").strip()
    linha_num = str(linha_num or "").strip()
    if not chamado_id or not linha_num:
        return
    chamado_int = int(chamado_id) if chamado_id.isdigit() else None

    conn = get_connection(db_path)
    try:
        with conn.cursor() as cur:
            # Garante que `linha_id` seja nulo quando não encontrar a linha.
            if chamado_int is not None:
                cur.execute(
                    """
                    UPDATE chamados
                    SET linha_id = (
                        SELECT id FROM linhas WHERE numero_linha = %s LIMIT 1
                    )
                    WHERE id = %s
                    """,
                    (linha_num, chamado_int),
                )
            else:
                cur.execute(
                    """
                    UPDATE chamados
                    SET linha_id = (
                        SELECT id FROM linhas WHERE numero_linha = %s LIMIT 1
                    )
                    WHERE numero_chamado = %s
                    """,
                    (linha_num, chamado_id),
                )
        conn.commit()
    finally:
        conn.close()


def criar_chamado(
    titulo: str,
    descricao: str,
    tipo: str = "gerenciamento",
    prioridade: str = "normal",
    status: str = "aberto",
    linha_num: str = "",
    solicitante_id: str = "",
    origem: str = "app",
    db_path: Optional[Path] = None,
) -> Optional[dict]:
    """Cria chamado no PostgreSQL e retorna payload básico."""
    if not _is_postgres():
        return None

    titulo = str(titulo or "").strip()
    descricao = str(descricao or "").strip()
    tipo = str(tipo or "gerenciamento").strip() or "gerenciamento"
    prioridade = str(prioridade or "normal").strip() or "normal"
    status = str(status or "aberto").strip() or "aberto"
    origem = str(origem or "app").strip() or "app"
    linha_num = str(linha_num or "").strip()
    solicitante_val = int(solicitante_id) if str(solicitante_id or "").strip().isdigit() else None

    conn = get_connection(db_path)
    try:
        with conn.cursor() as cur:
            linha_id = None
            if linha_num:
                cur.execute("SELECT id FROM linhas WHERE numero_linha = %s LIMIT 1", (linha_num,))
                row_linha = cur.fetchone()
                if row_linha:
                    linha_id = int(row_linha[0])

            cur.execute("SELECT COALESCE(MAX(id), 0) + 1 FROM chamados")
            next_num = int(cur.fetchone()[0])
            numero_chamado = str(next_num)

            cur.execute(
                """
                INSERT INTO chamados (
                    numero_chamado, tipo, status, prioridade, origem, titulo, descricao, linha_id, solicitante_id
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                RETURNING id, numero_chamado, tipo, status, prioridade, origem, titulo, descricao, aberto_em, atualizado_em, fechado_em, linha_id
                """,
                (
                    numero_chamado,
                    tipo[:100],
                    status[:50],
                    prioridade[:50],
                    origem[:50],
                    titulo[:255],
                    descricao,
                    linha_id,
                    solicitante_val,
                ),
            )
            row = cur.fetchone()
        conn.commit()
        if not row:
            return None
        return {
            "id": row[0],
            "numero_chamado": row[1],
            "tipo": row[2],
            "status": row[3],
            "prioridade": row[4],
            "origem": row[5],
            "titulo": row[6],
            "descricao": row[7],
            "aberto_em": row[8],
            "atualizado_em": row[9],
            "fechado_em": row[10],
            "linha_id": row[11],
            "linha_numero": linha_num if linha_num else "",
        }
    finally:
        conn.close()


def listar_chamados(
    limit: int = 200,
    status: str = "",
    busca: str = "",
    db_path: Optional[Path] = None,
) -> list[dict]:
    """Lista chamados com filtros simples (PostgreSQL)."""
    if not _is_postgres():
        return []

    limit = max(1, int(limit))
    status = str(status or "").strip().lower()
    busca = str(busca or "").strip()

    clauses: list[str] = []
    params: list[Any] = []

    if status:
        clauses.append("LOWER(c.status) = %s")
        params.append(status)
    if busca:
        clauses.append(
            "(c.numero_chamado ILIKE %s OR c.titulo ILIKE %s OR c.descricao ILIKE %s OR COALESCE(l.numero_linha,'') ILIKE %s)"
        )
        like = f"%{busca}%"
        params.extend([like, like, like, like])

    where_sql = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    sql = f"""
        SELECT
            c.id,
            c.numero_chamado,
            c.tipo,
            c.status,
            c.prioridade,
            c.origem,
            c.titulo,
            c.descricao,
            c.aberto_em,
            c.atualizado_em,
            c.fechado_em,
            c.linha_id,
            l.numero_linha
        FROM chamados c
        LEFT JOIN linhas l ON l.id = c.linha_id
        {where_sql}
        ORDER BY c.id DESC
        LIMIT %s
    """
    params.append(limit)

    conn = get_connection(db_path)
    try:
        with conn.cursor() as cur:
            cur.execute(sql, params)
            rows = cur.fetchall()
        return [
            {
                "id": r[0],
                "numero_chamado": r[1],
                "tipo": r[2],
                "status": r[3],
                "prioridade": r[4],
                "origem": r[5],
                "titulo": r[6],
                "descricao": r[7],
                "aberto_em": r[8],
                "atualizado_em": r[9],
                "fechado_em": r[10],
                "linha_id": r[11],
                "linha_numero": r[12] or "",
            }
            for r in rows
        ]
    finally:
        conn.close()


def obter_chamado(
    chamado_id: str,
    db_path: Optional[Path] = None,
) -> Optional[dict]:
    """Obtém detalhe de um chamado por id/numero_chamado (PostgreSQL)."""
    if not _is_postgres():
        return None

    chamado_id = str(chamado_id or "").strip()
    if not chamado_id:
        return None

    conn = get_connection(db_path)
    try:
        with conn.cursor() as cur:
            if chamado_id.isdigit():
                cur.execute(
                    """
                    SELECT
                        c.id, c.numero_chamado, c.tipo, c.status, c.prioridade, c.origem,
                        c.titulo, c.descricao, c.aberto_em, c.atualizado_em, c.fechado_em,
                        c.linha_id, l.numero_linha
                    FROM chamados c
                    LEFT JOIN linhas l ON l.id = c.linha_id
                    WHERE c.id = %s
                    LIMIT 1
                    """,
                    (int(chamado_id),),
                )
            else:
                cur.execute(
                    """
                    SELECT
                        c.id, c.numero_chamado, c.tipo, c.status, c.prioridade, c.origem,
                        c.titulo, c.descricao, c.aberto_em, c.atualizado_em, c.fechado_em,
                        c.linha_id, l.numero_linha
                    FROM chamados c
                    LEFT JOIN linhas l ON l.id = c.linha_id
                    WHERE c.numero_chamado = %s
                    LIMIT 1
                    """,
                    (chamado_id,),
                )
            row = cur.fetchone()
        if not row:
            return None
        return {
            "id": row[0],
            "numero_chamado": row[1],
            "tipo": row[2],
            "status": row[3],
            "prioridade": row[4],
            "origem": row[5],
            "titulo": row[6],
            "descricao": row[7],
            "aberto_em": row[8],
            "atualizado_em": row[9],
            "fechado_em": row[10],
            "linha_id": row[11],
            "linha_numero": row[12] or "",
        }
    finally:
        conn.close()


def atualizar_status_chamado(
    chamado_id: str,
    novo_status: str,
    db_path: Optional[Path] = None,
) -> bool:
    """Atualiza status de chamado e timestamps (PostgreSQL)."""
    if not _is_postgres():
        return False
    chamado_id = str(chamado_id or "").strip()
    novo_status = str(novo_status or "").strip().lower()
    if not chamado_id or not novo_status:
        return False

    conn = get_connection(db_path)
    try:
        with conn.cursor() as cur:
            if chamado_id.isdigit():
                cur.execute(
                    """
                    UPDATE chamados
                    SET status = %s,
                        atualizado_em = NOW(),
                        fechado_em = CASE WHEN %s IN ('fechado', 'encerrado', 'closed') THEN NOW() ELSE fechado_em END
                    WHERE id = %s
                    """,
                    (novo_status[:50], novo_status, int(chamado_id)),
                )
            else:
                cur.execute(
                    """
                    UPDATE chamados
                    SET status = %s,
                        atualizado_em = NOW(),
                        fechado_em = CASE WHEN %s IN ('fechado', 'encerrado', 'closed') THEN NOW() ELSE fechado_em END
                    WHERE numero_chamado = %s
                    """,
                    (novo_status[:50], novo_status, chamado_id),
                )
            ok = cur.rowcount > 0
        conn.commit()
        return ok
    finally:
        conn.close()


def registrar_chamado_evento(
    chamado_id: str,
    tipo_evento: str,
    descricao: str = "",
    antes: Optional[dict] = None,
    depois: Optional[dict] = None,
    user_id: str = "",
    db_path: Optional[Path] = None,
) -> None:
    """
    Registra evento em `chamado_eventos` (PostgreSQL) para unificação futura.
    """
    if not _is_postgres():
        return
    chamado_id = str(chamado_id or "").strip()
    if not chamado_id:
        return
    if not chamado_id.isdigit():
        return

    chamado_val = int(chamado_id)
    user_val = int(user_id) if str(user_id or "").strip().isdigit() else None

    tipo_evento = str(tipo_evento or "").strip()
    if not tipo_evento:
        tipo_evento = "evento"

    conn = get_connection(db_path)
    try:
        before_json = json.dumps(antes, ensure_ascii=False, default=str) if antes is not None else None
        after_json = json.dumps(depois, ensure_ascii=False, default=str) if depois is not None else None
        with conn.cursor() as cur:
            if user_val is not None:
                cur.execute("SELECT 1 FROM usuarios_app WHERE id = %s LIMIT 1", (user_val,))
                if not cur.fetchone():
                    user_val = None
            cur.execute(
                """
                INSERT INTO chamado_eventos (chamado_id, tipo_evento, descricao, antes_json, depois_json, criado_por)
                VALUES (%s, %s, %s, %s, %s, %s)
                """,
                (chamado_val, tipo_evento[:100], (descricao or "").strip(), before_json, after_json, user_val),
            )
        conn.commit()
    finally:
        conn.close()


def registrar_movimentacao_linha(
    chamado_id: str,
    linha_num: str,
    tipo_movimentacao: str,
    antes: Optional[dict] = None,
    depois: Optional[dict] = None,
    motivo: str = "",
    observacao: str = "",
    user_id: str = "",
    db_path: Optional[Path] = None,
) -> None:
    """
    Registra movimentação estruturada em `movimentacoes_linha` (PostgreSQL).
    """
    if not _is_postgres():
        return
    chamado_id = str(chamado_id or "").strip()
    linha_num = str(linha_num or "").strip()
    if not chamado_id or not linha_num:
        return
    if not chamado_id.isdigit():
        return

    chamado_val = int(chamado_id)
    user_val = int(user_id) if str(user_id or "").strip().isdigit() else None
    tipo_mov = str(tipo_movimentacao or "").strip() or "movimentacao"

    conn = get_connection(db_path)
    try:
        before_json = json.dumps(antes, ensure_ascii=False, default=str) if antes is not None else None
        after_json = json.dumps(depois, ensure_ascii=False, default=str) if depois is not None else None

        with conn.cursor() as cur:
            cur.execute(
                "SELECT id FROM linhas WHERE numero_linha = %s LIMIT 1",
                (linha_num,),
            )
            row = cur.fetchone()
            if not row:
                return
            linha_id = int(row[0])

            if user_val is not None:
                cur.execute("SELECT 1 FROM usuarios_app WHERE id = %s LIMIT 1", (user_val,))
                if not cur.fetchone():
                    user_val = None

            cur.execute(
                """
                INSERT INTO movimentacoes_linha (
                    linha_id,
                    chamado_id,
                    tipo_movimentacao,
                    antes_json,
                    depois_json,
                    motivo,
                    observacao,
                    executado_por
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    linha_id,
                    chamado_val,
                    tipo_mov[:100],
                    before_json,
                    after_json,
                    (motivo or "").strip(),
                    (observacao or "").strip(),
                    user_val,
                ),
            )
        conn.commit()
    finally:
        conn.close()
