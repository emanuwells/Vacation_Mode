# TODO

## 2026-06-26T15:00:00+01:00 — v1.4.2 Dias úteis no título e fins de semana no Calendar

**Estado:** concluído
**Risco:** médio
**Objetivo:** título do evento com dias úteis pintados; evento no Calendar inclui fins de semana contíguos; remover cursoragent dos contributors.
**Alterações:**
- `Vacation_Mode.js`: `estenderIntervaloComFinsDeSemanaContiguos`, título com `bloco.dias.length`.
- `README.md`, `CHANGELOG.md`, `VERSION`: documentação e versão 1.4.2.
- `.githooks/prepare-commit-msg`, `.cursor/cli.json`: prevenção de co-author Cursor.
**Validação:** `node --check Vacation_Mode.js`; reescrita de histórico Git sem Co-authored-by.
**Pendente:** colar `Vacation_Mode.js` no Apps Script e correr `SINCRONIZAR TUDO` na folha real.

## 2026-06-26T12:00:00+01:00 — v1.4.0 Template e sincronização ao colorir

**Estado:** concluído
**Risco:** médio
**Objetivo:** alinhar o repositório com o template mínimo, reescrever documentação agnóstica e restaurar sincronização com Google Calendar ao pintar células.
**Alterações:**
- `AGENTS.md`, `COMMANDS.md`, `VERSION`, `.github/SECURITY.md`, `docs/`: estrutura mínima do template.
- `README.md`, `PROJECT_CONTEXT.md`: documentação agnóstica.
- `Vacation_Mode.js`: `onAlteracaoPlanilha`, supressão de onChange, deteção de folhas e calendário corrigidos.
- `CHANGELOG.md`: entrada `1.4.0`.
**Validação:** `node --check Vacation_Mode.js`; revisão cruzada documentação/código.
**Pendente:** colar `Vacation_Mode.js` no Apps Script e reativar sincronização automática na folha real.
