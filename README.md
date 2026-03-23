# Gerenciamento de Telefones

Painel para controle de linhas telefônicas por equipe, operando com PostgreSQL como backend principal e preparado para o módulo de chamados no mesmo banco.

## Estrutura do Projeto

```
Planilhas Telefones/
├── app.py                 # Aplicação principal (Streamlit)
├── run.py                 # Ponto de entrada
├── requirements.txt
├── src/
│   ├── core/
│   │   └── config.py      # Configurações do projeto
│   ├── utils/
│   │   ├── text.py        # Normalização de texto
│   │   └── validators.py  # Validações
│   └── db/
│       ├── schema.sql            # Schema legado SQLite
│       ├── schema_postgres.sql   # Schema alvo PostgreSQL
│       └── repository.py         # Acesso ao banco
├── scripts/
│   ├── init_postgres.py
│   ├── migrate_sqlite_to_postgres.py
│   ├── aplicar_equipes_alimento.py
│   └── atualizar_regras_equipes.py
├── data/
│   └── db/               # Banco SQLite legado/local
└── doc/                  # Regras e documentação
    ├── equipe_regras.csv
    ├── equipes_alimento.csv
    └── equipes_medicamento.csv
```

## Como usar

### Estado atual da arquitetura

O projeto opera com:

- **PostgreSQL** como banco principal
- **backend como unica fonte de dados**
- **Gerenciamento + Chamados** convergindo para o mesmo sistema

As planilhas nao fazem mais parte do fluxo operacional da aplicacao.

### Opção rápida

- **Windows:** Dê duplo clique em **`ativador.bat`** — instala dependências e inicia o sistema.
- **Linux/Servidor:** Execute **`./ativador.sh`** (antes: `chmod +x ativador.sh`). O app escuta em todas as interfaces (0.0.0.0) para acesso remoto.

### Login

Na primeira execução, crie o usuário administrador na tela de login. Depois, administradores podem criar outros usuários em **Config** -> **Gerenciar usuarios**.

### Integração com Chamados (auditoria)

Se abrir o Gerenciamento com um parâmetro de chamado na URL, o ID é vinculado automaticamente nos eventos de auditoria:

- `?chamado_id=123`
- `?id_chamado=123`
- `?ticket_id=123`
- `?chamado=123`
- `?linha=11999999999`
- `?segmento_chamado=Alimento`
- `?equipe_chamado=Equipe%20A`
- `?return_url=https://sistema.exemplo/chamados/123`

Exemplo:

`http://localhost:8501/?chamado_id=123&linha=11999999999&segmento_chamado=Alimento&return_url=https://sistema.exemplo/chamados/123`

### Instalação manual

#### 1. Instalar dependências

```bash
pip install -r requirements.txt
```

#### 2. Configurar o banco principal

Defina a variável de ambiente `DATABASE_URL` apontando para o PostgreSQL:

```bash
DATABASE_URL=postgresql://usuario:senha@host:5432/gerenciamento_telefones
```

Forma mais simples no projeto:

1. copie `.env.example` para `.env`
2. troque `SUA_SENHA` pela senha real do PostgreSQL

Exemplo:

```bash
DATABASE_URL=postgresql://postgres:SUA_SENHA@localhost:5432/gerenciamento_telefones
```

Schema alvo do banco:

```bash
src/db/schema_postgres.sql
```

Para criar as tabelas no PostgreSQL:

```bash
python -m scripts.init_postgres
```

Para migrar os dados atuais do SQLite legado para o PostgreSQL:

```bash
python -m scripts.migrate_sqlite_to_postgres
```

#### 3. Executar o painel

```bash
streamlit run app.py
```

Ou:

```bash
python run.py
```

#### 4. Direção do projeto

Direção aprovada para as próximas fases:

- usar somente o banco do backend;
- unificar chamados e gerenciamento;
- centralizar auditoria e movimentações no mesmo banco.

## Modos e Segmentos

- **Linhas ativas**: registros ativos no banco
- **Linhas desativadas**: registros desativados no banco
- **Segmentos**: Alimento, Medicamento, Promotores

## Arquivos de configuração (doc/)

- `equipe_regras.csv`: regras de equipes, gestores e supervisores
- `equipes_alimento.csv`: mapeamento Alimento
- `equipes_medicamento.csv`: mapeamento Medicamento

## Instalação em outro computador ou servidor

1. Copie a pasta completa do projeto.
2. **Windows:** Execute **`ativador.bat`**.
3. **Linux/Servidor:** Execute `chmod +x ativador.sh` e depois `./ativador.sh`.
4. O sistema será instalado e iniciado. Em servidor, acesse pelo IP: `http://IP_DO_SERVIDOR:8501`.
