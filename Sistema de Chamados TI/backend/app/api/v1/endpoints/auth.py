import base64
import hashlib
import hmac
import json
import os
import secrets
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, Optional

from fastapi import APIRouter, HTTPException, status
from pydantic import BaseModel, Field

try:
    import psycopg
except ImportError:  # pragma: no cover
    psycopg = None


router = APIRouter(tags=["auth"], prefix="/api/auth")


class SSOExchangeRequest(BaseModel):
    sso_code: str = Field(..., min_length=1)


class SSOExchangeResponse(BaseModel):
    access_token: str
    refresh_token: str
    token_type: str = "bearer"
    token_expires_in_seconds: int
    refresh_token_expires_in_seconds: int


def _require_database_url() -> str:
    db_url = os.environ.get("DATABASE_URL", "").strip()
    if not db_url:
        raise RuntimeError("DATABASE_URL nao configurado no ambiente.")
    return db_url


def _require_jwt_secret() -> str:
    # Fallback para o mesmo segredo já usado pelo Streamlit (evita precisar criar mais variáveis).
    secret = os.environ.get("JWT_SECRET") or os.environ.get("COOKIES_PASSWORD")
    if not secret:
        raise RuntimeError("JWT_SECRET nao configurado e COOKIES_PASSWORD nao existe.")
    return str(secret)


def _b64url_encode(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).decode("utf-8").rstrip("=")


def _create_hs256_jwt(payload: Dict[str, Any], secret: str) -> str:
    header = {"alg": "HS256", "typ": "JWT"}
    header_b64 = _b64url_encode(json.dumps(header, separators=(",", ":")).encode("utf-8"))
    payload_b64 = _b64url_encode(json.dumps(payload, separators=(",", ":")).encode("utf-8"))
    signing_input = f"{header_b64}.{payload_b64}".encode("utf-8")
    signature = hmac.new(secret.encode("utf-8"), signing_input, hashlib.sha256).digest()
    sig_b64 = _b64url_encode(signature)
    return f"{header_b64}.{payload_b64}.{sig_b64}"


def _jwt_payload_common(*, user_id: int, username: str, is_admin: bool, token_type: str, now: datetime) -> Dict[str, Any]:
    return {
        "iss": "planilhas-telefones",
        "sub": str(user_id),
        "username": username,
        "is_admin": bool(is_admin),
        "type": token_type,
        "iat": int(now.timestamp()),
        # jti ajuda a reduzir risco caso alguém reutilize tokens antigas.
        "jti": secrets.token_hex(16),
    }


@router.post("/sso-exchange", response_model=SSOExchangeResponse)
def sso_exchange(body: SSOExchangeRequest) -> SSOExchangeResponse:
    """
    Troca 1 uso de `sso_code` por `access_token` + `refresh_token` (JWT HS256).

    Regras:
    - valida `sso_code` no PostgreSQL unificado (tabela `sso_codes`)
    - garante 1 uso via UPDATE atômico (usado_em NULL e expira_em > now())
    - marca `usado_em = now()` no mesmo UPDATE
    """
    if psycopg is None:  # pragma: no cover
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="psycopg nao instalado no ambiente do backend.",
        )

    sso_code = (body.sso_code or "").strip()
    if not sso_code:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="sso_code invalido.")

    db_url = _require_database_url()
    jwt_secret = _require_jwt_secret()

    access_ttl_seconds = int(os.environ.get("JWT_ACCESS_EXPIRES_SECONDS", str(15 * 60)))
    refresh_ttl_seconds = int(os.environ.get("JWT_REFRESH_EXPIRES_SECONDS", str(30 * 24 * 60 * 60)))

    now = datetime.now(timezone.utc)
    access_exp = now + timedelta(seconds=access_ttl_seconds)
    refresh_exp = now + timedelta(seconds=refresh_ttl_seconds)

    update_sql = """
        UPDATE sso_codes
        SET usado_em = NOW()
        WHERE code = %s
          AND usado_em IS NULL
          AND expira_em > NOW()
        RETURNING usuario_app_id
    """

    with psycopg.connect(db_url) as conn:
        try:
            with conn.cursor() as cur:
                cur.execute(update_sql, (sso_code,))
                row = cur.fetchone()
                if not row:
                    raise HTTPException(
                        status_code=status.HTTP_401_UNAUTHORIZED,
                        detail="SSO code invalido ou expirado (ou ja utilizado).",
                    )
                usuario_app_id = int(row[0])

                cur.execute(
                    """
                    SELECT id, username, is_admin
                    FROM usuarios_app
                    WHERE id = %s
                    """,
                    (usuario_app_id,),
                )
                user = cur.fetchone()
                if not user:
                    raise HTTPException(
                        status_code=status.HTTP_401_UNAUTHORIZED,
                        detail="Usuario nao encontrado para o SSO code.",
                    )

                user_id = int(user[0])
                username = str(user[1] or "").strip()
                is_admin = bool(user[2])

                # Constrói tokens com expiração.
                access_payload = _jwt_payload_common(
                    user_id=user_id,
                    username=username,
                    is_admin=is_admin,
                    token_type="access",
                    now=now,
                )
                access_payload["exp"] = int(access_exp.timestamp())

                refresh_payload = _jwt_payload_common(
                    user_id=user_id,
                    username=username,
                    is_admin=is_admin,
                    token_type="refresh",
                    now=now,
                )
                refresh_payload["exp"] = int(refresh_exp.timestamp())

                access_token = _create_hs256_jwt(access_payload, jwt_secret)
                refresh_token = _create_hs256_jwt(refresh_payload, jwt_secret)

                return SSOExchangeResponse(
                    access_token=access_token,
                    refresh_token=refresh_token,
                    token_type="bearer",
                    token_expires_in_seconds=access_ttl_seconds,
                    refresh_token_expires_in_seconds=refresh_ttl_seconds,
                )
        except HTTPException:
            # Não mascarar erros de validação.
            raise
        except Exception as exc:
            # Ajuda a depurar problemas de schema (tabela/coluna) sem perder contexto.
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=f"Erro ao processar troca de SSO: {type(exc).__name__}: {exc}",
            )

