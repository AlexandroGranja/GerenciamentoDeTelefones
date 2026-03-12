"""Funções de normalização e manipulação de texto."""

import unicodedata
from typing import Any


def normalize_text(value: Any) -> str:
    """Normaliza texto: remove acentos e converte para minúsculas."""
    if value is None:
        return ""
    text = str(value).strip()
    text = "".join(
        ch for ch in unicodedata.normalize("NFD", text)
        if unicodedata.category(ch) != "Mn"
    )
    return text.lower()


def title_case_safe(value: str) -> str:
    """Aplica title case preservando siglas (até 4 letras)."""
    tokens = [t for t in value.replace("/", " / ").split(" ") if t != ""]
    if not tokens:
        return "Sem Equipe"
    out: list[str] = []
    for token in tokens:
        if token == "/":
            out.append(token)
        elif token.isupper() and len(token) <= 4:
            out.append(token)
        else:
            out.append(token.capitalize())
    return " ".join(out).replace(" / ", "/")


def normalize_team_key(value: Any) -> str:
    """Chave normalizada para equipe."""
    text = normalize_text(value)
    return " ".join(text.split())


def digits_only(value: Any) -> str:
    """Extrai apenas dígitos do valor."""
    if value is None:
        return ""
    return "".join(ch for ch in str(value) if ch.isdigit())
