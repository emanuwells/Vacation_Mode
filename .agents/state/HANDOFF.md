# HANDOFF

- **Última atualização:** 2026-07-26T23:12:00+01:00
- **Estado:** concluído — WELLS 0.5.0 em `master` e `origin/master`
- **Branch:** `master` (`b67f882`); feature `chore/wells-ai-toolkit` também no remoto
- **Objetivo:** migração toolkit e alinhamento Git concluídos

## Estado útil

- **Concluído:** runtime `.agents/` 0.5.0; docs alinhadas; v1.4.4; merge FF para `master`; push OK
- **Em curso:** N/A
- **Bloqueios/riscos:** Apps Script em produção atualiza-se por cópia manual (só cabeçalho de versão mudou)
- **Ficheiros relevantes:** `.agents/AGENTS.md`, `VERSION`, `CHANGELOG.md`
- **Validações executadas:** `node --check Vacation_Mode.js`; `node .agents/tools/validate-project.mjs`; `HEAD == origin/master`
- **Próximo passo exato:** opcional — colar `Vacation_Mode.js` no Apps Script se quiser o cabeçalho 1.4.4 na folha real
