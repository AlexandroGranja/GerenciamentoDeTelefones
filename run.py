#!/usr/bin/env python
"""Ponto de entrada da aplicação Gerenciamento de Telefones."""

import subprocess
import sys
from pathlib import Path

if __name__ == "__main__":
    root = Path(__file__).resolve().parent
    sys.exit(subprocess.run([sys.executable, "-m", "streamlit", "run", "app.py", "--server.headless", "true"], cwd=root).returncode)
