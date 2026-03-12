-- Schema do banco de dados - Gerenciamento de Telefones
-- Tabela principal com dados processados (denormalizado para performance)

CREATE TABLE IF NOT EXISTS linhas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo TEXT,
    nome TEXT,
    equipe TEXT,
    equipe_padrao TEXT,
    grupo_equipe TEXT,
    tipo_equipe TEXT,
    localidade TEXT,
    data_troca TEXT,
    data_retorno TEXT,
    data_ocorrencia TEXT,
    data_solicitacao_tbs TEXT,
    gestor TEXT,
    supervisor TEXT,
    segmento TEXT,
    papel TEXT,
    linha TEXT NOT NULL,
    email TEXT,
    gerenciamento TEXT,
    imei_a TEXT,
    imei_b TEXT,
    marca TEXT,
    chip TEXT,
    aparelho TEXT,
    modelo TEXT,
    setor TEXT,
    cargo TEXT,
    desconto TEXT,
    perfil TEXT,
    empresa TEXT,
    ativo TEXT,
    numero_serie TEXT,
    patrimonio TEXT,
    operadora TEXT,
    nome_guerra TEXT,
    motivo TEXT,
    observacao TEXT,
    aba TEXT,
    modo TEXT DEFAULT 'ativas',
    criado_em TEXT DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_linhas_modo ON linhas(modo);
CREATE INDEX IF NOT EXISTS idx_linhas_segmento ON linhas(segmento);
CREATE INDEX IF NOT EXISTS idx_linhas_equipe_padrao ON linhas(equipe_padrao);
CREATE INDEX IF NOT EXISTS idx_linhas_linha ON linhas(linha);

-- Tabela de relação (linhas ativas) para verificação
CREATE TABLE IF NOT EXISTS relacao_ativas (
    linha TEXT PRIMARY KEY,
    atualizado_em TEXT DEFAULT (datetime('now'))
);

-- Tabela de usuários para login
CREATE TABLE IF NOT EXISTS usuarios (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    salt TEXT NOT NULL,
    is_admin INTEGER DEFAULT 0,
    criado_em TEXT DEFAULT (datetime('now'))
);
CREATE UNIQUE INDEX IF NOT EXISTS idx_usuarios_username ON usuarios(username);

-- Tabela de sessões para login persistente (cookies)
CREATE TABLE IF NOT EXISTS sessoes (
    token TEXT PRIMARY KEY,
    username TEXT NOT NULL,
    criado_em TEXT DEFAULT (datetime('now')),
    expira_em TEXT NOT NULL,
    FOREIGN KEY (username) REFERENCES usuarios(username)
);
CREATE INDEX IF NOT EXISTS idx_sessoes_username ON sessoes(username);
CREATE INDEX IF NOT EXISTS idx_sessoes_expira ON sessoes(expira_em);

-- Tabela de auditoria para rastrear alterações do sistema
CREATE TABLE IF NOT EXISTS audit_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    acao TEXT NOT NULL,
    entidade TEXT NOT NULL,
    chave_registro TEXT,
    chamado_id TEXT,
    antes_json TEXT,
    depois_json TEXT,
    detalhes TEXT,
    user_id TEXT,
    username TEXT,
    origem TEXT DEFAULT 'app',
    criado_em TEXT DEFAULT (datetime('now'))
);
CREATE INDEX IF NOT EXISTS idx_audit_criado_em ON audit_log(criado_em);
CREATE INDEX IF NOT EXISTS idx_audit_username ON audit_log(username);
CREATE INDEX IF NOT EXISTS idx_audit_entidade ON audit_log(entidade);
