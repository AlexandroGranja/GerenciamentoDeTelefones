# Arquitetura Alvo - Sistema Unificado

## Decisao

O projeto deixa de usar planilhas como fonte de verdade.

A partir desta nova fase:

- o banco principal sera PostgreSQL;
- o backend sera a unica camada de leitura e escrita;
- Gerenciamento de Telefones e Sistema de Chamados passam a evoluir como um sistema unificado.

## Objetivo

Ter um unico sistema para operar:

- linhas;
- usuarios;
- aparelhos;
- chips;
- chamados;
- movimentacoes;
- auditoria.

## Principios

1. Planilha nao e mais banco de dados.
2. Toda alteracao operacional precisa gerar auditoria.
3. Chamado e linha precisam estar ligados no mesmo modelo.
4. Regra de negocio precisa morar no backend, nao na planilha.
5. O app deve consumir apenas o banco principal.

## Modulos do sistema

### 1. Identidade e acesso

- `usuarios_app`
- `perfis`
- `usuarios_app_perfis`
- `sessoes`

### 2. Estrutura organizacional

- `segmentos`
- `equipes`
- `colaboradores`

### 3. Ativos de telefonia

- `linhas`
- `aparelhos`
- `chips`

### 4. Operacao de chamados

- `chamados`
- `chamado_eventos`

### 5. Historico operacional

- `movimentacoes_linha`
- `auditoria`

## Regras de negocio principais

### Linha vaga

Uma linha sera considerada vaga quando:

- existir numero de linha;
- nao existir usuario vinculado;
- ou o snapshot do usuario estiver marcado como `VAGO`.

No banco unificado, isso deve ser representado preferencialmente por:

- `colaborador_id = NULL`
- `nome_usuario_snapshot = 'VAGO'`
- `codigo_usuario_snapshot = 'VAGO'`
- `status = 'vaga'` quando fizer sentido operacional

### Abertura de chamado

Quando um chamado impactar uma linha:

- o chamado deve apontar para `linha_id`;
- a movimentacao deve ser registrada em `movimentacoes_linha`;
- a auditoria deve registrar antes/depois;
- o historico do chamado deve mostrar os eventos relacionados.

### Mudancas operacionais

Eventos como:

- desativacao;
- ativacao;
- envio para manutencao;
- roubo e perda;
- troca de equipe;
- troca de segmento;
- marcar linha como vaga;

devem registrar:

- quem executou;
- quando executou;
- motivo;
- estado anterior;
- estado final.

## Ordem recomendada de implementacao

### Etapa 1 - Fundacao do banco

1. Criar schema PostgreSQL unificado.
2. Definir `DATABASE_URL`.
3. Preparar conexao do backend para PostgreSQL.

### Etapa 2 - Migracao do Gerenciamento

1. Remover fallback para planilhas.
2. Remover sincronizacao planilha -> banco.
3. Fazer a aplicacao ler e gravar somente pelo backend.

### Etapa 3 - Integracao de chamados

1. Criar tabela `chamados`.
2. Criar `chamado_eventos`.
3. Relacionar `chamados` com `linhas`.
4. Levar o fluxo de abertura/atualizacao para dentro do mesmo sistema.

### Etapa 4 - Unificacao operacional

1. Uma interface unica.
2. Um login unico.
3. Uma auditoria unica.
4. Uma trilha de movimentacao unica.

## Entregavel desta fase

Nesta fase inicial, o projeto passa a ter:

- schema alvo PostgreSQL em `src/db/schema_postgres.sql`;
- configuracao preparada para `DATABASE_URL`;
- documentacao da arquitetura alvo;
- roadmap atualizado para a migracao DB-only e unificacao dos sistemas.
