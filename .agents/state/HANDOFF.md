# HANDOFF

- **Última atualização:** 2026-08-20T12:00:00+01:00
- **Estado:** concluído — v1.5.1 em `master`, ainda por colar na folha real
- **Branch:** `master`
- **Objetivo:** corrigir a sincronização automática do Calendar (regressão de 1.5.0) e mover o script para `src/`

## Estado útil

- **Concluído:** `sincronizarBlocosComDiferenca` volta a isolar cada bloco num `try/catch` (um bloco com erro não aborta os restantes; erro de quota continua a interromper de imediato); `Vacation_Mode.js` movido para `src/Vacation_Mode.js`; documentação e testes atualizados; v1.5.1.
- **Em curso:** N/A
- **Bloqueios/riscos:** o utilizador reportou que "SINCRONIZAR TUDO" falhava depois de ativar a sincronização automática (1.5.0); a causa mais provável identificada por revisão de código foi a falta de isolamento por bloco (corrigida agora). Não foi possível confirmar com o log real de Execuções do Apps Script nem com a folha real — a confirmação definitiva só é possível do lado do utilizador.
- **Ficheiros relevantes:** `src/Vacation_Mode.js`, `tests/triggers.test.js`, `tests/calendar.test.js`, `CHANGELOG.md`, `VERSION`
- **Validações executadas:** `node --check src/Vacation_Mode.js`; `node tests/triggers.test.js` (5/5); `node tests/calendar.test.js` (7/7, incluindo os 2 cenários novos de isolamento por bloco); `node .agents/tools/validate-project.mjs` (`ok: true`)
- **Próximo passo exato:** colar `src/Vacation_Mode.js` no Apps Script (substitui o código atual por completo, não só a diferença), reativar "Ativar Sincronização Automática" para reinstalar os triggers e limpar bloqueios de quota antigos, e correr "SINCRONIZAR TUDO". Se ainda falhar, consultar "Execuções" no editor do Apps Script e reportar a mensagem de erro exata para diagnóstico dirigido.
