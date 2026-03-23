# Mapa de Melhorias - Gerenciamento de Telefones

## Objetivo

Evoluir o sistema de Gerenciamento de Telefones por etapas, com foco em:

- rastreabilidade real das alteracoes;
- integracao progressiva com o Sistema de Chamados;
- seguranca e previsibilidade no login;
- usabilidade para o time operacional;
- robustez de dados e manutencao simples.

---

## Estado Atual

### Ja implementado

1. Banco SQLite com persistencia atual do gerenciamento.
2. Login local com sessao persistente por cookie.
3. Login unico dentro do Gerenciamento (uma sessao ativa por usuario).
4. Auditoria de acoes principais no banco (`audit_log`).
5. Historico simplificado no app para admin.
6. Registro detalhado de edicoes com campo, antes e depois.
7. Vinculo opcional com `chamado_id` via query string.
8. Historico mostrando:
   - quem editou;
   - chamado;
   - o que editou;
   - quando editou.
9. Painel operacional por modo e segmento.
10. Filtro de linhas vagas e regra funcional de `VAGO`.

### Limitacoes atuais

1. O sistema ainda carrega partes do fluxo pensando em SQLite como banco atual.
2. Ainda existe codigo legado para planilhas e sincronizacao planilha -> banco.
3. O sistema de chamados ainda nao esta dentro do mesmo backend.
4. O login unico ainda vale so no Gerenciamento, nao entre os dois sistemas.

---

## Roadmap por Fases

### Fase 1 - Auditoria e rastreabilidade fina

Status: **concluida**

Objetivo:
Deixar toda alteracao importante rastreavel, legivel e vinculavel ao chamado.

Entregas:

1. Manter auditoria humana e clara no app.
2. Registrar alteracoes por campo (`antes -> depois`) de forma consistente.
3. Mostrar `chamado_id` no historico.
4. Melhorar ainda mais a granularidade do historico quando necessario.

Pendencias opcionais desta fase:

1. Opcional: transformar cada alteracao de campo em um evento individual.
2. Criar filtro por linha no historico.

---

### Fase 2 - Integracao operacional com chamados

Status: **concluida**

Objetivo:
Fazer o Gerenciamento trabalhar junto com o Sistema de Chamados no fluxo diario.

Entregas:

1. Abrir o Gerenciamento ja com contexto do chamado.
2. Permitir localizar automaticamente linha/equipe a partir do chamado.
3. Manter `chamado_id` em todo fluxo relevante.
4. Padronizar retorno visual para o operador apos salvar alteracao vinculada ao chamado.

Sugestoes tecnicas:

1. Aceitar `chamado_id`, `linha`, `segmento` e `equipe` por URL.
2. Destacar na interface quando o usuario estiver atuando em contexto de chamado.
3. Criar acao de retorno ao sistema de chamados apos salvar.

---

### Fase 3 - Seguranca, identidade e acesso

Status: **planejada**

Objetivo:
Unificar identidade e tornar o acesso mais seguro e previsivel.

Entregas:

1. Definir identificador unico entre sistemas (preferencia: email corporativo).
2. Mapear usuarios do Chamados e do Gerenciamento.
3. Evoluir de login local para identidade unificada.
4. Preparar terreno para SSO real.

Depois disso:

1. Escolher provedor OIDC (`Keycloak`, `Authentik`, `Azure AD` ou equivalente).
2. Implementar autenticacao nos dois sistemas.
3. Unificar logout e sessao.

---

### Fase 4 - Confiabilidade de edicao

Status: **concluida**

Objetivo:
Evitar perda de dados e sobrescrita silenciosa.

Entregas:

1. Detectar conflito de edicao concorrente.
2. Avisar quando uma linha foi alterada por outro usuario.
3. Salvar somente registros realmente modificados.
4. Melhorar mensagens de erro e recuperacao.

---

### Fase 5 - UX e operacao

Status: **em andamento**

Objetivo:
Deixar o sistema mais rapido e claro para uso diario.

Entregas:

1. Dashboard operacional com indicadores uteis.
2. Filtros mais fortes e possivelmente filtros salvos.
3. Fluxos guiados para manutencao, roubo/perda e desativacao.
4. Configuracao mais clara por blocos.
5. Validacoes melhores de campos criticos (linha, IMEI, datas, email).

---

### Fase 6 - Observabilidade e resiliência

Status: **planejada**

Objetivo:
Dar visibilidade tecnica e reduzir risco operacional.

Entregas:

1. Logs tecnicos separados da auditoria do usuario.
2. Painel de saude do sistema.
3. Backup automatizado do banco.
4. Reprocessamento manual de operacoes com erro.

---

### Fase 7 - Banco definitivo e sistema unificado

Status: **iniciada**

Objetivo:

Migrar o projeto para um banco definitivo em PostgreSQL e unir Gerenciamento + Chamados em um unico sistema.

Entregas:

1. Definir PostgreSQL como banco principal.
2. Remover uso de planilhas como fonte de verdade.
3. Criar schema unificado para:
   - usuarios;
   - linhas;
   - aparelhos;
   - chips;
   - chamados;
   - movimentacoes;
   - auditoria.
4. Mover o fluxo de chamados para dentro do mesmo backend.
5. Preparar unificacao de login, permissoes e historico.

Artefatos desta fase:

1. `src/db/schema_postgres.sql`
2. `doc/ARQUITETURA_SISTEMA_UNIFICADO.md`

---

## Prioridade Recomendada

### Fazer agora

1. Fase 7 - Banco definitivo e sistema unificado.
2. Concluir a retirada de planilhas do app.
3. Consolidar a Fase 5 - UX e operacao sobre banco unico.

### Fazer depois

1. Fase 3 - Identidade unificada e SSO.
2. Fase 6 - Observabilidade e resiliência.

---

## Proximo Passo Sugerido

### Parte 1

Iniciar a **migracao DB-only**:

1. remover fallback para planilhas no app;
2. remover sincronizacao planilha -> banco da operacao diaria;
3. preparar o backend para usar `DATABASE_URL` como fonte principal.

### Parte 2

Depois disso, iniciar a unificacao com chamados:

1. criar tabelas de `chamados` e `chamado_eventos`;
2. ligar `chamados` com `linhas`;
3. mover o fluxo do sistema de chamados para dentro do gerenciamento.
