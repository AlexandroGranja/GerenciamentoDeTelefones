# Gerenciamento de Telefones

Painel para controle de linhas telefônicas por equipe, com suporte a banco de dados SQLite.

## Estrutura do Projeto

```
Planilhas Telefones/
├── app.py                 # Aplicação principal (Streamlit)
├── run.py                 # Ponto de entrada
├── requirements.txt
├── src/
│   ├── core/
│   │   └── config.py      # Configurações e constantes
│   ├── utils/
│   │   ├── text.py        # Normalização de texto
│   │   └── validators.py  # Validações
│   └── db/
│       ├── schema.sql    # Schema do banco
│       └── repository.py # Acesso ao banco
├── scripts/
│   ├── sync_db.py        # Sincroniza planilhas → banco
│   ├── aplicar_equipes_alimento.py
│   ├── atualizar_regras_equipes.py
│   └── listar_inativas.py
├── data/
│   ├── planilhas/        # Planilhas Excel
│   │   ├── Telefones11.25_SomenteAtivas.xlsx
│   │   ├── Telefones11.25.xlsx
│   │   └── Relação Linhas Prosper270226.xlsx
│   └── db/               # Banco SQLite gerado
└── doc/                  # Regras e documentação
    ├── equipe_regras.csv
    ├── equipes_alimento.csv
    └── equipes_medicamento.csv
```

## Como usar

### Opção rápida

- **Windows:** Dê duplo clique em **`ativador.bat`** — instala dependências, sincroniza o banco e inicia o sistema.
- **Linux/Servidor:** Execute **`./ativador.sh`** (antes: `chmod +x ativador.sh`). O app escuta em todas as interfaces (0.0.0.0) para acesso remoto.

### Login

Na primeira execução, crie o usuário administrador na tela de login. Depois, administradores podem criar outros usuários em **⚙ Config** → **Gerenciar usuários**.

### Integração com Chamados (auditoria)

Se abrir o Gerenciamento com um parâmetro de chamado na URL, o ID é vinculado automaticamente nos eventos de auditoria:

- `?chamado_id=123`
- `?id_chamado=123`
- `?ticket_id=123`
- `?chamado=123`

Exemplo:

`http://localhost:8501/?chamado_id=123`

### Instalação manual

#### 1. Instalar dependências

```bash
pip install -r requirements.txt
```

#### 2. Sincronizar o banco (recomendado)

Para não depender das planilhas em toda execução, sincronize uma vez:

```bash
python -m scripts.sync_db
```

Isso cria `data/db/gerenciamento_telefones.db` com os dados processados.

#### 3. Executar o painel

```bash
streamlit run app.py
```

Ou:

```bash
python run.py
```

#### 4. Modo de dados

Em **⚙ Config** você pode:
- **Usar banco de dados** (padrão): leitura rápida do SQLite
- **Usar planilhas**: lê direto do Excel (mais lento)
- **Sincronizar banco**: atualiza o DB a partir das planilhas

## Modos e Segmentos

- **Linhas ativas**: Telefones11.25_SomenteAtivas.xlsx
- **Linhas desativadas**: Telefones11.25.xlsx (exceto linhas na Relação)
- **Segmentos**: Alimento, Medicamento, Promotores

## Arquivos de configuração (doc/)

- `equipe_regras.csv`: regras de equipes, gestores e supervisores
- `equipes_alimento.csv`: mapeamento Alimento
- `equipes_medicamento.csv`: mapeamento Medicamento

## Instalação em outro computador ou servidor

1. Copie a pasta completa do projeto (incluindo `data/planilhas/` e `doc/` com os arquivos).
2. **Windows:** Execute **`ativador.bat`**.
3. **Linux/Servidor:** Execute `chmod +x ativador.sh` e depois `./ativador.sh`.
4. O sistema será instalado e iniciado. Em servidor, acesse pelo IP: `http://IP_DO_SERVIDOR:8501`.
