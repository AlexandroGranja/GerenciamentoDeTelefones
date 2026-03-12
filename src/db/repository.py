"""Repositório para acesso ao banco de dados."""

import hashlib
import json
import secrets
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

import pandas as pd

from src.core.config import get_db_path


def _hash_password(password: str, salt: str) -> str:
    """Gera hash da senha com salt."""
    return hashlib.sha256((salt + password).encode("utf-8")).hexdigest()


def criar_usuario(username: str, password: str, is_admin: bool = False, db_path: Optional[Path] = None) -> bool:
    """Cria novo usuário. Retorna True se sucesso."""
    salt = secrets.token_hex(16)
    password_hash = _hash_password(password, salt)
    conn = get_connection(db_path)
    try:
        conn.execute(
            "INSERT INTO usuarios (username, password_hash, salt, is_admin) VALUES (?, ?, ?, ?)",
            (username.strip().lower(), password_hash, salt, 1 if is_admin else 0),
        )
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()


def verificar_login(username: str, password: str, db_path: Optional[Path] = None) -> Optional[dict]:
    """
    Verifica credenciais. Retorna dict com username, is_admin ou None.
    """
    conn = get_connection(db_path)
    try:
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
        cur = conn.execute("DELETE FROM usuarios WHERE username = ?", (username.strip().lower(),))
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()


def tem_usuarios(db_path: Optional[Path] = None) -> bool:
    """Verifica se existe pelo menos um usuário."""
    conn = get_connection(db_path)
    try:
        return conn.execute("SELECT COUNT(*) FROM usuarios").fetchone()[0] > 0
    finally:
        conn.close()


def obter_usuario_por_username(username: str, db_path: Optional[Path] = None) -> Optional[dict]:
    """Retorna usuário por username (sem senha). Para restaurar sessão."""
    conn = get_connection(db_path)
    try:
        row = conn.execute(
            "SELECT username, is_admin FROM usuarios WHERE username = ?",
            (username.strip().lower(),),
        ).fetchone()
        if not row:
            return None
        return {"username": row[0], "is_admin": bool(row[1])}
    finally:
        conn.close()


def criar_sessao(token: str, username: str, dias_validade: int = 30, db_path: Optional[Path] = None) -> None:
    """Cria sessão para login persistente (login único por usuário)."""
    expira = (datetime.now() + timedelta(days=dias_validade)).strftime("%Y-%m-%d %H:%M:%S")
    conn = get_connection(db_path)
    try:
        # Login único: ao autenticar, invalida sessões antigas do mesmo usuário.
        conn.execute("DELETE FROM sessoes WHERE username = ?", (username.strip().lower(),))
        conn.execute(
            "INSERT OR REPLACE INTO sessoes (token, username, expira_em) VALUES (?, ?, ?)",
            (token, username.strip().lower(), expira),
        )
        conn.commit()
    finally:
        conn.close()


def validar_sessao(token: str, db_path: Optional[Path] = None) -> Optional[dict]:
    """Valida token de sessão. Retorna dados do usuário ou None."""
    conn = get_connection(db_path)
    try:
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
        conn.execute("DELETE FROM sessoes WHERE token = ?", (token,))
        conn.commit()
    finally:
        conn.close()


def get_connection(db_path: Optional[Path] = None) -> sqlite3.Connection:
    """Retorna conexão SQLite."""
    path = db_path or get_db_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path))
    conn.row_factory = sqlite3.Row
    return conn


def init_db(db_path: Optional[Path] = None) -> None:
    """Inicializa o schema do banco."""
    schema_path = Path(__file__).parent / "schema.sql"
    conn = get_connection(db_path)
    try:
        conn.executescript(schema_path.read_text(encoding="utf-8"))
        # Migração leve para bancos já existentes.
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
    """Salva DataFrame na tabela linhas. Remove registros antigos do modo e insere novos."""
    col_map = {
        "Codigo": "codigo", "Nome": "nome", "Equipe": "equipe",
        "EquipePadrao": "equipe_padrao", "GrupoEquipe": "grupo_equipe",
        "TipoEquipe": "tipo_equipe", "Localidade": "localidade",
        "Data da Troca": "data_troca", "Data Retorno": "data_retorno",
        "Data Ocorrência": "data_ocorrencia", "Data Solicitação TBS": "data_solicitacao_tbs",
        "Gestor": "gestor", "Supervisor": "supervisor", "Segmento": "segmento",
        "Papel": "papel", "Linha": "linha", "E-mail": "email",
        "Gerenciamento": "gerenciamento", "IMEI A": "imei_a", "IMEI B": "imei_b",
        "CHIP": "chip", "Marca": "marca", "Aparelho": "aparelho", "Modelo": "modelo",
        "Setor": "setor", "Cargo": "cargo", "Desconto": "desconto",
        "Perfil": "perfil", "Empresa": "empresa", "Ativo": "ativo",
        "Numero de Serie": "numero_serie", "Patrimonio": "patrimonio", "Operadora": "operadora",
        "Nome de Guerra": "nome_guerra", "Motivo": "motivo", "Observação": "observacao", "Aba": "aba",
    }
    conn = get_connection(db_path)
    try:
        conn.execute("DELETE FROM linhas WHERE modo = ?", (modo,))
        available = [c for c in col_map if c in df.columns]
        df_export = df[available].copy()
        df_export = df_export.rename(columns=col_map)
        df_export["modo"] = modo
        df_export.to_sql("linhas", conn, if_exists="append", index=False)
        conn.commit()
        return len(df_export)
    finally:
        conn.close()


def load_linhas(modo: str = "ativas", db_path: Optional[Path] = None) -> pd.DataFrame:
    """Carrega linhas do banco. Retorna DataFrame vazio se não houver dados."""
    rev_map = {
        "codigo": "Codigo", "nome": "Nome", "equipe": "Equipe",
        "equipe_padrao": "EquipePadrao", "grupo_equipe": "GrupoEquipe",
        "tipo_equipe": "TipoEquipe", "localidade": "Localidade",
        "data_troca": "Data da Troca", "data_retorno": "Data Retorno",
        "data_ocorrencia": "Data Ocorrência", "data_solicitacao_tbs": "Data Solicitação TBS",
        "gestor": "Gestor", "supervisor": "Supervisor", "segmento": "Segmento",
        "papel": "Papel", "linha": "Linha", "email": "E-mail",
        "gerenciamento": "Gerenciamento", "imei_a": "IMEI A", "imei_b": "IMEI B",
        "chip": "CHIP", "marca": "Marca", "aparelho": "Aparelho", "modelo": "Modelo",
        "setor": "Setor", "cargo": "Cargo", "desconto": "Desconto",
        "perfil": "Perfil", "empresa": "Empresa", "ativo": "Ativo",
        "numero_serie": "Numero de Serie", "patrimonio": "Patrimonio", "operadora": "Operadora",
        "nome_guerra": "Nome de Guerra", "motivo": "Motivo", "observacao": "Observação", "aba": "Aba",
    }
    conn = get_connection(db_path)
    try:
        df = pd.read_sql_query(
            "SELECT * FROM linhas WHERE modo = ?",
            conn,
            params=(modo,),
        )
        if df.empty:
            return pd.DataFrame()
        df = df.drop(columns=[c for c in df.columns if c in ("id", "criado_em", "modo")], errors="ignore")
        df = df.rename(columns={c: rev_map[c] for c in df.columns if c in rev_map})
        return df
    finally:
        conn.close()


def save_relacao_ativas(linhas: set[str], db_path: Optional[Path] = None) -> int:
    """Salva conjunto de linhas ativas na relação."""
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
                json.dumps(antes, ensure_ascii=False) if antes is not None else None,
                json.dumps(depois, ensure_ascii=False) if depois is not None else None,
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
    db_path: Optional[Path] = None,
) -> list[dict]:
    """Lista eventos de auditoria mais recentes (com filtros opcionais)."""
    conn = get_connection(db_path)
    try:
        clauses: list[str] = []
        params: list[object] = []
        if username.strip():
            clauses.append("username = ?")
            params.append(username.strip())
        if acao.strip():
            clauses.append("acao = ?")
            params.append(acao.strip())
        if entidade.strip():
            clauses.append("entidade = ?")
            params.append(entidade.strip())
        where_sql = f"WHERE {' AND '.join(clauses)}" if clauses else ""
        sql = f"""
            SELECT id, acao, entidade, chave_registro, chamado_id, antes_json, depois_json, detalhes, user_id, username, origem, criado_em
            FROM audit_log
            {where_sql}
            ORDER BY id DESC
            LIMIT ?
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
