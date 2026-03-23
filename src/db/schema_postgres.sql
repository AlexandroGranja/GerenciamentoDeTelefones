-- Schema alvo PostgreSQL - Sistema unificado de Gerenciamento + Chamados
-- Fonte de verdade unica da operacao.

CREATE TABLE IF NOT EXISTS perfis (
    id BIGSERIAL PRIMARY KEY,
    codigo VARCHAR(50) NOT NULL UNIQUE,
    nome VARCHAR(100) NOT NULL,
    descricao TEXT,
    criado_em TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS usuarios_app (
    id BIGSERIAL PRIMARY KEY,
    username VARCHAR(150) NOT NULL UNIQUE,
    email VARCHAR(255),
    nome_exibicao VARCHAR(255),
    password_hash TEXT,
    salt TEXT,
    auth_provider VARCHAR(50) NOT NULL DEFAULT 'local',
    is_admin BOOLEAN NOT NULL DEFAULT FALSE,
    ativo BOOLEAN NOT NULL DEFAULT TRUE,
    criado_em TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    atualizado_em TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS usuarios_app_perfis (
    usuario_id BIGINT NOT NULL REFERENCES usuarios_app(id) ON DELETE CASCADE,
    perfil_id BIGINT NOT NULL REFERENCES perfis(id) ON DELETE CASCADE,
    PRIMARY KEY (usuario_id, perfil_id)
);

CREATE TABLE IF NOT EXISTS sessoes (
    token VARCHAR(255) PRIMARY KEY,
    usuario_id BIGINT NOT NULL REFERENCES usuarios_app(id) ON DELETE CASCADE,
    criado_em TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    expira_em TIMESTAMPTZ NOT NULL
);

-- Códigos de SSO (1 uso) para redirecionar do Gerenciamento (Streamlit)
-- para o sistema React externo (Chamados).
CREATE TABLE IF NOT EXISTS sso_codes (
    id BIGSERIAL PRIMARY KEY,
    code VARCHAR(255) NOT NULL UNIQUE,
    usuario_id BIGINT NOT NULL REFERENCES usuarios_app(id) ON DELETE CASCADE,
    criado_em TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    expira_em TIMESTAMPTZ NOT NULL,
    usado_em TIMESTAMPTZ
);

CREATE INDEX IF NOT EXISTS idx_sso_codes_usuario_id ON sso_codes(usuario_id);
CREATE INDEX IF NOT EXISTS idx_sso_codes_expira_em ON sso_codes(expira_em);

CREATE TABLE IF NOT EXISTS segmentos (
    id BIGSERIAL PRIMARY KEY,
    codigo VARCHAR(50) NOT NULL UNIQUE,
    nome VARCHAR(100) NOT NULL UNIQUE,
    ativo BOOLEAN NOT NULL DEFAULT TRUE,
    criado_em TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS equipes (
    id BIGSERIAL PRIMARY KEY,
    segmento_id BIGINT REFERENCES segmentos(id),
    codigo VARCHAR(50),
    nome VARCHAR(150) NOT NULL,
    tipo_equipe VARCHAR(100),
    grupo_equipe VARCHAR(150),
    gestor_nome VARCHAR(255),
    supervisor_nome VARCHAR(255),
    localidade VARCHAR(255),
    ativo BOOLEAN NOT NULL DEFAULT TRUE,
    criado_em TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    atualizado_em TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_equipes_segmento_id ON equipes(segmento_id);
CREATE INDEX IF NOT EXISTS idx_equipes_nome ON equipes(nome);

CREATE TABLE IF NOT EXISTS colaboradores (
    id BIGSERIAL PRIMARY KEY,
    codigo VARCHAR(100),
    nome VARCHAR(255) NOT NULL,
    nome_guerra VARCHAR(255),
    email VARCHAR(255),
    equipe_id BIGINT REFERENCES equipes(id),
    segmento_id BIGINT REFERENCES segmentos(id),
    cargo VARCHAR(255),
    localidade VARCHAR(255),
    gestor_nome VARCHAR(255),
    supervisor_nome VARCHAR(255),
    ativo BOOLEAN NOT NULL DEFAULT TRUE,
    criado_em TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    atualizado_em TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_colaboradores_codigo ON colaboradores(codigo);
CREATE INDEX IF NOT EXISTS idx_colaboradores_nome ON colaboradores(nome);
CREATE INDEX IF NOT EXISTS idx_colaboradores_equipe_id ON colaboradores(equipe_id);

CREATE TABLE IF NOT EXISTS aparelhos (
    id BIGSERIAL PRIMARY KEY,
    patrimonio VARCHAR(100),
    numero_serie VARCHAR(100),
    marca VARCHAR(100),
    modelo VARCHAR(150),
    imei_a VARCHAR(50),
    imei_b VARCHAR(50),
    status VARCHAR(50) NOT NULL DEFAULT 'ativo',
    observacao TEXT,
    criado_em TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    atualizado_em TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_aparelhos_imei_a ON aparelhos(imei_a);
CREATE INDEX IF NOT EXISTS idx_aparelhos_imei_b ON aparelhos(imei_b);

CREATE TABLE IF NOT EXISTS chips (
    id BIGSERIAL PRIMARY KEY,
    numero_chip VARCHAR(100),
    operadora VARCHAR(100),
    status VARCHAR(50) NOT NULL DEFAULT 'ativo',
    criado_em TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    atualizado_em TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS linhas (
    id BIGSERIAL PRIMARY KEY,
    numero_linha VARCHAR(30) NOT NULL UNIQUE,
    linha VARCHAR(30),
    status VARCHAR(50) NOT NULL DEFAULT 'ativa',
    modo_operacao VARCHAR(30) NOT NULL DEFAULT 'ativas',
    modo VARCHAR(30) NOT NULL DEFAULT 'ativas',
    colaborador_id BIGINT REFERENCES colaboradores(id),
    nome_usuario_snapshot VARCHAR(255),
    codigo_usuario_snapshot VARCHAR(100),
    email_snapshot VARCHAR(255),
    ordem_manual INTEGER,
    codigo VARCHAR(100),
    nome VARCHAR(255),
    equipe VARCHAR(150),
    equipe_padrao VARCHAR(150),
    grupo_equipe VARCHAR(150),
    tipo_equipe VARCHAR(100),
    localidade VARCHAR(255),
    gestor VARCHAR(255),
    supervisor VARCHAR(255),
    segmento VARCHAR(100),
    papel VARCHAR(100),
    email VARCHAR(255),
    equipe_id BIGINT REFERENCES equipes(id),
    segmento_id BIGINT REFERENCES segmentos(id),
    aparelho_id BIGINT REFERENCES aparelhos(id),
    chip_id BIGINT REFERENCES chips(id),
    gerenciamento TEXT,
    data_troca TEXT,
    data_retorno TEXT,
    data_ocorrencia TEXT,
    data_solicitacao_tbs TEXT,
    motivo TEXT,
    observacao TEXT,
    setor VARCHAR(255),
    cargo VARCHAR(255),
    desconto VARCHAR(100),
    perfil VARCHAR(100),
    empresa VARCHAR(100),
    ativo_texto VARCHAR(50),
    ativo VARCHAR(50),
    aba_origem VARCHAR(100),
    aba VARCHAR(100),
    marca VARCHAR(100),
    imei_a VARCHAR(50),
    imei_b VARCHAR(50),
    chip VARCHAR(100),
    aparelho VARCHAR(150),
    modelo VARCHAR(150),
    numero_serie VARCHAR(100),
    patrimonio VARCHAR(100),
    operadora VARCHAR(100),
    nome_guerra VARCHAR(255),
    origem_registro VARCHAR(50) NOT NULL DEFAULT 'sistema',
    criado_em TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    atualizado_em TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_linhas_status ON linhas(status);
CREATE INDEX IF NOT EXISTS idx_linhas_modo_operacao ON linhas(modo_operacao);
CREATE INDEX IF NOT EXISTS idx_linhas_modo ON linhas(modo);
CREATE INDEX IF NOT EXISTS idx_linhas_linha ON linhas(linha);
CREATE INDEX IF NOT EXISTS idx_linhas_segmento ON linhas(segmento);
CREATE INDEX IF NOT EXISTS idx_linhas_equipe_padrao ON linhas(equipe_padrao);
CREATE INDEX IF NOT EXISTS idx_linhas_segmento_id ON linhas(segmento_id);
CREATE INDEX IF NOT EXISTS idx_linhas_equipe_id ON linhas(equipe_id);
CREATE INDEX IF NOT EXISTS idx_linhas_colaborador_id ON linhas(colaborador_id);

CREATE TABLE IF NOT EXISTS chamados (
    id BIGSERIAL PRIMARY KEY,
    numero_chamado VARCHAR(100) NOT NULL UNIQUE,
    tipo VARCHAR(100) NOT NULL,
    status VARCHAR(50) NOT NULL DEFAULT 'aberto',
    prioridade VARCHAR(50) NOT NULL DEFAULT 'normal',
    origem VARCHAR(50) NOT NULL DEFAULT 'gerenciamento',
    titulo VARCHAR(255),
    descricao TEXT,
    -- Campos compatíveis com a UI do sistema de chamados (React)
    category VARCHAR(100),
    subcategory VARCHAR(100),
    location VARCHAR(100),
    equipment_info TEXT,
    internal_notes TEXT,
    linha_id BIGINT REFERENCES linhas(id),
    colaborador_id BIGINT REFERENCES colaboradores(id),
    solicitante_id BIGINT REFERENCES usuarios_app(id),
    responsavel_id BIGINT REFERENCES usuarios_app(id),
    aberto_em TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    atualizado_em TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    fechado_em TIMESTAMPTZ
);

CREATE INDEX IF NOT EXISTS idx_chamados_status ON chamados(status);
CREATE INDEX IF NOT EXISTS idx_chamados_linha_id ON chamados(linha_id);
CREATE INDEX IF NOT EXISTS idx_chamados_responsavel_id ON chamados(responsavel_id);

CREATE TABLE IF NOT EXISTS chamado_eventos (
    id BIGSERIAL PRIMARY KEY,
    chamado_id BIGINT NOT NULL REFERENCES chamados(id) ON DELETE CASCADE,
    tipo_evento VARCHAR(100) NOT NULL,
    descricao TEXT,
    antes_json JSONB,
    depois_json JSONB,
    criado_por BIGINT REFERENCES usuarios_app(id),
    criado_em TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_chamado_eventos_chamado_id ON chamado_eventos(chamado_id);
CREATE INDEX IF NOT EXISTS idx_chamado_eventos_tipo ON chamado_eventos(tipo_evento);

CREATE TABLE IF NOT EXISTS movimentacoes_linha (
    id BIGSERIAL PRIMARY KEY,
    linha_id BIGINT NOT NULL REFERENCES linhas(id) ON DELETE CASCADE,
    chamado_id BIGINT REFERENCES chamados(id),
    tipo_movimentacao VARCHAR(100) NOT NULL,
    de_status VARCHAR(50),
    para_status VARCHAR(50),
    de_equipe_id BIGINT REFERENCES equipes(id),
    para_equipe_id BIGINT REFERENCES equipes(id),
    de_segmento_id BIGINT REFERENCES segmentos(id),
    para_segmento_id BIGINT REFERENCES segmentos(id),
    antes_json JSONB,
    depois_json JSONB,
    motivo TEXT,
    observacao TEXT,
    executado_por BIGINT REFERENCES usuarios_app(id),
    criado_em TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_movimentacoes_linha_linha_id ON movimentacoes_linha(linha_id);
CREATE INDEX IF NOT EXISTS idx_movimentacoes_linha_chamado_id ON movimentacoes_linha(chamado_id);

CREATE TABLE IF NOT EXISTS auditoria (
    id BIGSERIAL PRIMARY KEY,
    acao VARCHAR(100) NOT NULL,
    entidade VARCHAR(100) NOT NULL,
    entidade_id VARCHAR(100),
    chamado_id BIGINT REFERENCES chamados(id),
    linha_id BIGINT REFERENCES linhas(id),
    antes_json JSONB,
    depois_json JSONB,
    detalhes TEXT,
    user_id BIGINT REFERENCES usuarios_app(id),
    username VARCHAR(150),
    origem VARCHAR(50) NOT NULL DEFAULT 'app',
    criado_em TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_auditoria_criado_em ON auditoria(criado_em DESC);
CREATE INDEX IF NOT EXISTS idx_auditoria_chamado_id ON auditoria(chamado_id);
CREATE INDEX IF NOT EXISTS idx_auditoria_linha_id ON auditoria(linha_id);
CREATE INDEX IF NOT EXISTS idx_auditoria_username ON auditoria(username);

INSERT INTO perfis (codigo, nome, descricao)
VALUES
    ('admin', 'Administrador', 'Acesso total ao sistema unificado'),
    ('operador', 'Operador', 'Opera linhas, aparelhos e chamados'),
    ('gestor', 'Gestor', 'Consulta indicadores e acompanha equipes')
ON CONFLICT (codigo) DO NOTHING;

INSERT INTO segmentos (codigo, nome)
VALUES
    ('alimento', 'Alimento'),
    ('medicamento', 'Medicamento'),
    ('promotores', 'Promotores'),
    ('internos', 'Internos'),
    ('manutencao', 'Manutenção'),
    ('roubo_perda', 'Roubo e Perda')
ON CONFLICT (codigo) DO NOTHING;
