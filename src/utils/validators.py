"""Funções de validação."""

from typing import Any

from src.utils.text import digits_only


def is_valid_phone(value: Any) -> bool:
    """Verifica se o valor é um número de telefone válido (10-13 dígitos)."""
    d = digits_only(value)
    return 10 <= len(d) <= 13
