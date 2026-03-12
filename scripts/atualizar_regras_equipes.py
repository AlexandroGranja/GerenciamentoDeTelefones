"""
Script para preencher eh_equipe e equipe_pai no equipe_regras.csv.
Equipes reais: CONSUMO *, ESPECIAL CONSUMO.
Localidades: mapeadas para sua equipe pai conforme estrutura da planilha.
"""
import pandas as pd
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DOC_DIR = ROOT / "doc"
RULES = DOC_DIR / "equipe_regras.csv"
if not RULES.exists():
    print("Arquivo equipe_regras.csv não encontrado.")
    exit(1)

df = pd.read_csv(RULES, dtype=str).fillna("")

# Colunas novas
if "eh_equipe" not in df.columns:
    df["eh_equipe"] = "False"
if "equipe_pai" not in df.columns:
    df["equipe_pai"] = ""

def norm(s):
    return str(s).strip().lower().replace(" ", " ")

# Equipes reais (são equipes, não localidades)
EQUIPES_REAIS = {
    "consumo baixada", "consumo oeste", "consumo zona norte", "consumo niteroi",
    "especial consumo",
}

# Localidades -> equipe pai (inferido da planilha)
LOCALIDADES_PAI = {
    # Consumo Baixada
    "nova iguacu i": "Consumo Baixada", "n iguacu": "Consumo Baixada", "n iguacu 2": "Consumo Baixada",
    "gramacho": "Consumo Baixada", "jardim iris": "Consumo Baixada", "jardim irís": "Consumo Baixada",
    "vila de cava": "Consumo Baixada", "olinda": "Consumo Baixada",
    # Consumo Oeste
    "bangu": "Consumo Oeste", "campo grande": "Consumo Oeste", "freguesia": "Consumo Oeste",
    "inhoaiba": "Consumo Oeste", "inhoaíba": "Consumo Oeste", "guaratiba": "Consumo Oeste",
    # Consumo Zona Norte
    "jacare": "Consumo Zona Norte", "jacaré": "Consumo Zona Norte",
    "encantado": "Consumo Zona Norte", "sao cristovao": "Consumo Zona Norte", "são cristóvão": "Consumo Zona Norte",
    "bonsucesso": "Consumo Zona Norte", "cavalcanti": "Consumo Zona Norte", "cavalcante": "Consumo Zona Norte",
    "galeao": "Consumo Zona Norte", "galeão": "Consumo Zona Norte", "centro": "Consumo Zona Norte",
    "itanhanga": "Consumo Zona Norte", "itanhangá": "Consumo Zona Norte",
    "flamengo": "Consumo Zona Norte", "copacabana": "Consumo Zona Norte",
    "vidigal": "Consumo Zona Norte", "tijuca": "Consumo Zona Norte",
    # Consumo Niteroi
    "sao goncalo": "Consumo Niteroi", "são gonçalo": "Consumo Niteroi",
    "regiao oceanica": "Consumo Niteroi", "região oceânica": "Consumo Niteroi",
    "niteroi 2": "Consumo Niteroi",
    # Equipe Especial (Rede Especial, Impulso, Rota Especial, Auto Serviço)
    "rede especial 2": "Equipe Especial", "rede especial 3": "Equipe Especial", "rede especial 4": "Equipe Especial",
    "rota especial 1": "Equipe Especial", "rota especial 2": "Equipe Especial", "rota especial 3": "Equipe Especial",
    "impulso 01": "Equipe Especial", "impulso 02": "Equipe Especial",
    "a s impulso 2": "Equipe Especial", "a s impulso 3": "Equipe Especial", "a s impulso 4": "Equipe Especial", "a s impulso 5": "Equipe Especial",
    "auto servico 01": "Equipe Especial", "auto servico 05": "Equipe Especial", "auto servico 02": "Equipe Especial",
    "auto servico 03": "Equipe Especial", "auto servico 04": "Equipe Especial",
}

for i, row in df.iterrows():
    key = norm(row.get("equipe_key", ""))
    if not key:
        continue
    if key in EQUIPES_REAIS:
        df.at[i, "eh_equipe"] = "True"
        df.at[i, "equipe_pai"] = ""
        if key == "especial consumo":
            df.at[i, "equipe_padrao"] = "Equipe Especial"
    elif key in LOCALIDADES_PAI:
        df.at[i, "eh_equipe"] = "False"
        df.at[i, "equipe_pai"] = LOCALIDADES_PAI[key]
    else:
        df.at[i, "eh_equipe"] = "True"
        df.at[i, "equipe_pai"] = ""

cols = list(df.columns)
if "eh_equipe" not in cols:
    cols.append("eh_equipe")
if "equipe_pai" not in cols:
    cols.append("equipe_pai")
df = df[[c for c in cols if c in df.columns]]
df.to_csv(RULES, index=False, encoding="utf-8-sig")
print(f"Atualizado: {len(df)} linhas. eh_equipe=True: {(df['eh_equipe']=='True').sum()}, com equipe_pai: {(df['equipe_pai']!='').sum()}")
