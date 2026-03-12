"""
Atualiza equipe_regras.csv com base em equipes_alimento.csv.
Mapeia cada valor da planilha (Bangu, CONSUMO OESTE, etc.) para:
- equipe_real: equipe para agrupamento (Consumo Oeste, Equipe Especial)
- gestor: Marcelo (Consumo) ou Marco (Especial)
- supervisor: da tabela de referência
"""
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
DOC_DIR = ROOT / "doc"

def norm(s):
    return str(s).strip().lower().replace("  ", " ").replace("ã","a").replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u").replace("ç","c")

REF = DOC_DIR / "equipes_alimento.csv"
RULES = DOC_DIR / "equipe_regras.csv"

if not REF.exists() or not RULES.exists():
    print("Arquivos nao encontrados.")
    exit(1)

ref = pd.read_csv(REF, dtype=str).fillna("")
rules = pd.read_csv(RULES, dtype=str).fillna("")

# Monta lookup: (equipe_real_norm, localidade_norm) -> gerente, supervisor
# E também: equipe_real_norm -> gerente, supervisor (para linha vazia localidade)
lookup = {}
for _, r in ref.iterrows():
    er = str(r["equipe_real"]).strip()
    loc = str(r["localidade"]).strip()
    er_norm = norm(er)
    loc_norm = norm(loc)
    key = (er_norm, loc_norm)
    lookup[key] = {"gerente": str(r["gerente"]).strip(), "supervisor": str(r["supervisor"]).strip()}
    if not loc_norm and er_norm:
        lookup[(er_norm, "")] = lookup[key]

# Também: localidade como chave direta -> equipe_real, gerente, supervisor
# Assim quando planilha tem "Bangu", buscamos localidade_norm="bangu"
loc_to_equipe = {}
for _, r in ref.iterrows():
    er = str(r["equipe_real"]).strip()
    loc = str(r["localidade"]).strip()
    if loc:
        loc_to_equipe[norm(loc)] = {"equipe_real": er, "gerente": str(r["gerente"]).strip(), "supervisor": str(r["supervisor"]).strip()}
    er_norm = norm(er)
    if er_norm and er_norm not in loc_to_equipe:
        loc_to_equipe[er_norm] = {"equipe_real": er, "gerente": str(r["gerente"]).strip(), "supervisor": str(r["supervisor"]).strip()}

# Consumo * e especial consumo
for k in ["consumo baixada","consumo oeste","consumo zona norte","consumo niteroi"]:
    loc_to_equipe[k] = {"equipe_real": "Consumo "+k.replace("consumo ","").title(), "gerente": "Marcelo Neves", "supervisor": ""}
loc_to_equipe["especial consumo"] = {"equipe_real": "Equipe Especial", "gerente": "Marco Antonio Neves Suzart", "supervisor": "Ricardo Cascao"}
for k in ["rede especial 2","rede especial 3","rede especial 4","rota especial 1","rota especial 2","rota especial 3","impulso 01","impulso 02","auto servico 05","a s impulso 2","a s impulso 3","a s impulso 4","a s impulso 5","auto servico 01"]:
    loc_to_equipe[k] = {"equipe_real": "Equipe Especial", "gerente": "Marco Antonio Neves Suzart", "supervisor": "Ricardo Cascao"}

# Atualiza regras
for i, row in rules.iterrows():
    key = norm(row.get("equipe_key",""))
    if key in loc_to_equipe:
        rules.at[i, "equipe_padrao"] = loc_to_equipe[key]["equipe_real"]
        rules.at[i, "gestor"] = loc_to_equipe[key]["gerente"]
        rules.at[i, "supervisor"] = loc_to_equipe[key]["supervisor"] or row.get("supervisor","")
        rules.at[i, "eh_equipe"] = "True" if key in ["consumo baixada","consumo oeste","consumo zona norte","consumo niteroi","especial consumo"] else "False"
        rules.at[i, "equipe_pai"] = loc_to_equipe[key]["equipe_real"] if rules.at[i, "eh_equipe"] == "False" else ""

rules.to_csv(RULES, index=False, encoding="utf-8-sig")
print("equipe_regras.csv atualizado com equipes_alimento.csv")
