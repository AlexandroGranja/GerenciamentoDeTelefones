# Mapa de Melhorias - Integração Chamados x Gerenciamento

## Objetivo

Garantir integração confiável entre:

- Sistema de Chamados TI
- Gerenciamento de Telefones

Com foco em rastreabilidade, login único e robustez operacional.

---

## Fases

### Fase 1 - Auditoria de Alterações (iniciada)

Status: **em andamento**

Entregas desta fase:

1. Tabela de auditoria no banco (`audit_log`).
2. Registro de ações críticas:
   - criação de linha
   - exclusão de linha
   - mover equipe/setor
   - ativar/desativar linha
   - envio para manutenção
   - salvar edições em lote
   - sincronização de banco
3. Visualização de histórico para admin no app.

Campos-chave previstos em auditoria:

- `acao`
- `entidade`
- `chave_registro`
- `antes_json`
- `depois_json`
- `detalhes`
- `username`
- `origem`
- `criado_em`

---

### Fase 2 - Identidade unificada

Status: **planejada**

Entregas:

1. Definir identificador único de usuário (email corporativo).
2. Alinhar usuários dos dois sistemas.
3. Opcional: tabela de mapeamento de identidade.

---

### Fase 3 - Login único (SSO)

Status: **planejada**

Entregas:

1. Escolha do provedor OIDC (Keycloak/AuthentiK/Azure AD).
2. Chamados autenticando via OIDC.
3. Gerenciamento autenticando via OIDC.
4. Logout e sessão unificados.

---

### Fase 4 - Observabilidade e resiliência

Status: **planejada**

Entregas:

1. Logs estruturados de integração.
2. Retry controlado para falhas transitórias.
3. Painel de saúde da integração.
4. Reprocessamento manual de operações com erro.

---

## Decisões já tomadas

1. Segmentos especiais (Manutenção e Roubo e Perda) exibem dados consolidados.
2. Importação da planilha foi expandida para capturar colunas específicas por segmento.
3. Fluxos de movimentação de linha já operam no banco com persistência.

---

## Próximos passos sugeridos

1. Revisar logs de auditoria com usuário chave.
2. Definir formato de exportação de auditoria (CSV/Excel).
3. Iniciar desenho técnico da Fase 2 (identidade unificada).
