#!/usr/bin/env python
"""
Cria o primeiro usuário administrador.
Execute uma vez para configurar o login inicial.

Uso: python -m scripts.criar_admin
"""

import getpass
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from src.db.repository import init_db, criar_usuario, tem_usuarios

def main():
    init_db()
    if tem_usuarios():
        print("Ja existem usuarios. Use o app para criar novos (como admin).")
        return
    print("Criando primeiro usuario administrador...")
    username = input("Usuario: ").strip()
    if not username:
        print("Usuario invalido.")
        sys.exit(1)
    password = getpass.getpass("Senha: ")
    if len(password) < 4:
        print("Senha deve ter pelo menos 4 caracteres.")
        sys.exit(1)
    if criar_usuario(username, password, is_admin=True):
        print(f"Usuario '{username}' criado com sucesso!")
    else:
        print("Erro ao criar usuario.")
        sys.exit(1)

if __name__ == "__main__":
    main()
