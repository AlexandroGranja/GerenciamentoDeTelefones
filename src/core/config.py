"""Configurações e constantes da aplicação."""

from pathlib import Path

# Diretórios base
BASE_DIR = Path(__file__).resolve().parent.parent.parent
DATA_DIR = BASE_DIR / "data"
PLANILHAS_DIR = DATA_DIR / "planilhas"
DOC_DIR = BASE_DIR / "doc"
DB_DIR = DATA_DIR / "db"

# Arquivos de dados
DEFAULT_FILE = "Telefones11.25_SomenteAtivas.xlsx"
DEFAULT_FILE_COMPLETO = "Telefones11.25.xlsx"
RELACAO_FILE = "Relação Linhas Prosper270226.xlsx"
RULES_FILE = "equipe_regras.csv"
EQUIPES_ALIMENTO_FILE = "equipes_alimento.csv"
EQUIPES_MEDICAMENTO_FILE = "equipes_medicamento.csv"

# Caminhos completos
def get_planilhas_path(filename: str) -> Path:
    """Retorna caminho para arquivo em data/planilhas."""
    return PLANILHAS_DIR / filename

def get_doc_path(filename: str) -> Path:
    """Retorna caminho para arquivo em doc/."""
    return DOC_DIR / filename

def get_rules_path() -> Path:
    """Retorna caminho da pasta doc com regras."""
    return DOC_DIR

# Compatibilidade
RAW_DIR = PLANILHAS_DIR
CONFIG_DIR = DOC_DIR

def get_db_path() -> Path:
    """Retorna caminho do banco SQLite."""
    DB_DIR.mkdir(parents=True, exist_ok=True)
    return DB_DIR / "gerenciamento_telefones.db"

# Abas e segmentos
ABAS_ALIMENTO = ["Nova Prosper"]
ABAS_MEDICAMENTO = ["Prosper Norte", "Prosper Sul"]
ABAS_FOCO = ["Prosper Norte", "Prosper Sul", "Nova Prosper", "Promotores", "Internos", "Troca de Aparelho", "Devolução Manutenção", "Roubo-Perda"]
ABAS_PROMOTORES = ["Promotores"]
ABAS_INTERNOS = ["Internos"]
ABAS_MANUTENCAO = ["Troca de Aparelho", "Devolução Manutenção", "Devolucao Manutencao"]
ABAS_ROUBO_PERDA = ["Roubo-Perda", "Roubo e Perda"]

EQUIPES_PROMOTORES = ["Promotores"]
EQUIPES_INTERNOS = ["Internos"]
EQUIPES_MANUTENCAO = ["Manutenção"]
EQUIPES_ROUBO_PERDA = ["Roubo e Perda"]
EQUIPES_ALIMENTO = [
    "Gerentes do Alimento",
    "Consumo Baixada",
    "Consumo Oeste",
    "Consumo Zona Norte",
    "Consumo Niteroi",
    "Equipe Especial",
    "Gerente Senior",
]
EQUIPES_MEDICAMENTO = [
    "Gerentes do Medicamento",
    "Prosper Norte",
    "Prosper Sul",
]
GESTORES_MEDICAMENTO = {
    "Prosper Norte": "Priscila Rangel Manhães",
    "Prosper Sul": "Gustavo Luis Dias De Armada",
}

# Colunas padrão das planilhas
DEFAULT_COLUMNS = [
    "Codigo", "Nome", "Nome de Guerra", "Equipe", "Linha", "E-mail", "Gerenciamento",
    "Data da Troca", "Data Retorno", "Data Ocorrência", "Data Solicitação TBS", "Motivo", "Observação",
    "IMEI A", "IMEI B", "CHIP", "Aparelho", "Modelo",
    "Marca",
    "Setor", "Cargo", "Desconto", "Perfil", "Empresa",
    "Ativo", "Numero de Serie", "Patrimonio", "Operadora",
]
