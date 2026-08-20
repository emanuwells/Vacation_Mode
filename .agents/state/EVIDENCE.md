# EVIDENCE

Registo factual das validações relevantes.

## 2026-08-20 — Isolamento por bloco + mover para `src/` / v1.5.1

- **Alteração:** `sincronizarBlocosComDiferenca` volta a isolar cada bloco num `try/catch` (regressão de 1.5.0); `Vacation_Mode.js` movido para `src/Vacation_Mode.js`
- **Comando/validação:** `node --check src/Vacation_Mode.js`
- **Resultado real:** exit code 0

- **Comando/validação:** `node tests/triggers.test.js`
- **Resultado real:** 5/5 `OK` (debounce, cadência diária, quota bloqueada)

- **Comando/validação:** `node tests/calendar.test.js`
- **Resultado real:** 7/7 `OK` (sem alterações, crescer bloco, despintar bloco, quota entre folhas, desbloqueio manual, isolamento de erro por bloco, interrupção imediata em erro de quota)

- **Comando/validação:** `node .agents/tools/validate-project.mjs`
- **Resultado real:** `{"ok":true,...}`
- **Limitações:** não confirma o comportamento na folha Google real nem reproduz o erro original reportado pelo utilizador (log de Execuções do Apps Script não disponível nesta sessão); a causa foi identificada por revisão de código, não por reprodução direta do erro.

## 2026-07-26 — Migração WELLS 0.5.0 / v1.4.4

- **Alteração:** instalação do runtime `.agents/`, remoção de legado IA, alinhamento de docs
- **Comando/validação:** `node --check Vacation_Mode.js`
- **Resultado real:** exit code 0
- **Limitações:** não executa o runtime Google Apps Script

- **Comando/validação:** `node .agents/tools/validate-project.mjs`
- **Resultado real:** `{"ok":true,"counts":{"skills":28,"roles":11,"policies":14,"workflows":7},"agentsWords":487}`; `manifest.json` / `toolkit-lock.json` versão `0.5.0`
- **Limitações:** valida estrutura WELLS, não a folha Google em produção

## Modelo

- **Data:** AAAA-MM-DD
- **Alteração:**
- **Comando/validação:**
- **Resultado real:**
- **Limitações:**
