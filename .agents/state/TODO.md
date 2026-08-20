# TODO

## 2026-08-20T12:00:00+01:00 — v1.5.1 Isolamento por bloco + `src/`

**Estado:** concluído (código); pendente confirmação do utilizador na folha real
**Risco:** médio (regressão numa correção recente)
**Objetivo:** corrigir "SINCRONIZAR TUDO" a falhar depois de ativar a sincronização automática (1.5.0); mover `Vacation_Mode.js` para `src/`.
**Alterações:**
- `src/Vacation_Mode.js` (movido de `Vacation_Mode.js`): `sincronizarBlocosComDiferenca` volta a isolar cada bloco num `try/catch` (regressão de 1.5.0, que tinha perdido a resiliência por bloco do código anterior); `REGEX_QUOTA_CALENDAR` partilhado; erro de quota continua a propagar-se e a interromper de imediato.
- `tests/`, `README.md`, `PROJECT_CONTEXT.md`, `COMMANDS.md`, `docs/ROOT_STRUCTURE.md`: caminhos atualizados para `src/Vacation_Mode.js`.
**Validação:** `node --check src/Vacation_Mode.js`; `node tests/triggers.test.js` (5/5); `node tests/calendar.test.js` (7/7); `node .agents/tools/validate-project.mjs` (`ok: true`).
**Pendente:** colar `src/Vacation_Mode.js` no Apps Script, reativar sincronização automática e confirmar "SINCRONIZAR TUDO" na folha real; se persistir, obter a mensagem exata do log de Execuções.

## 2026-07-26T23:09:45+01:00 — v1.4.4 WELLS toolkit 0.5.0

**Estado:** concluído
**Risco:** baixo
**Objetivo:** migrar o sistema de agentes para WELLS 0.5.0, alinhar documentação e integrar em `master`.
**Alterações:**
- `.agents/` (runtime 0.5.0); removidos `AGENTS.md`, `tasks/`, `docs/ai/`, `tools/ai-adapters/`.
- `README.md`, `PROJECT_CONTEXT.md`, `COMMANDS.md`, `docs/ROOT_STRUCTURE.md`, `CHANGELOG.md`, `VERSION`.
- `CONTRIBUTING.md`, `SECURITY.md` na raiz.
**Validação:** `node --check Vacation_Mode.js`; `node .agents/tools/validate-project.mjs` (`ok: true`, toolkit `0.5.0`).
**Pendente:** se a folha real ainda tiver cabeçalho antigo, colar `Vacation_Mode.js` no Apps Script (sem mudança funcional).

## 2026-06-26T16:00:00+01:00 — v1.4.3 Raiz limpa e sem cursoragent

**Estado:** concluído
**Risco:** baixo
**Objetivo:** remover `.cursor/` e `.githooks/` da raiz; prevenir `cursoragent` nos contributors; alinhar Git com remoto.
**Alterações:**
- Removidos `.cursor/`, `.githooks/`, `scripts/install-git-hooks.ps1`.
- Documentação e regras de attribution atualizadas (hoje em `.agents/adapters/`).
**Validação:** estrutura conforme template; histórico sem Co-authored-by Cursor.

## 2026-06-26T15:00:00+01:00 — v1.4.2 Dias úteis no título e fins de semana no Calendar

**Estado:** concluído
**Risco:** médio
**Objetivo:** título do evento com dias úteis pintados; evento no Calendar inclui fins de semana contíguos; remover cursoragent dos contributors.
**Alterações:**
- `Vacation_Mode.js`: `estenderIntervaloComFinsDeSemanaContiguos`, título com `bloco.dias.length`.
- `README.md`, `CHANGELOG.md`, `VERSION`: documentação e versão 1.4.2.
**Validação:** `node --check Vacation_Mode.js`; reescrita de histórico Git sem Co-authored-by.
**Pendente:** colar `Vacation_Mode.js` no Apps Script e correr `SINCRONIZAR TUDO` na folha real.

## 2026-06-26T12:00:00+01:00 — v1.4.0 Template e sincronização ao colorir

**Estado:** concluído
**Risco:** médio
**Objetivo:** alinhar o repositório com o template mínimo, reescrever documentação agnóstica e restaurar sincronização com Google Calendar ao pintar células.
**Alterações:**
- Estrutura mínima do template; documentação agnóstica.
- `Vacation_Mode.js`: `onAlteracaoPlanilha`, supressão de onChange, deteção de folhas e calendário corrigidos.
- `CHANGELOG.md`: entrada `1.4.0`.
**Validação:** `node --check Vacation_Mode.js`; revisão cruzada documentação/código.
**Pendente:** colar `Vacation_Mode.js` no Apps Script e reativar sincronização automática na folha real.
