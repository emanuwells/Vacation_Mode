# EVIDENCE

Registo factual das validações relevantes.

## 2026-08-20 — Resumo final de "SINCRONIZAR TUDO" sem falsos positivos / v1.5.3

- **Alteração:** `sincronizarComCalendar` devolve `{status, resultado?}`; `sincronizarTudo` compõe a notificação final a partir dos resultados reais de cada folha
- **Evidência de origem:** segundo log real de Execuções enviado pelo utilizador, já com a correção de idioma (1.5.2) a funcionar corretamente: "Calendário 2026" bloqueado pela quota real do Google (comportamento correto); "Calendário 2025" sincronizado com 0 alterações (eventos já existentes, comportamento correto); mas a notificação final de "SINCRONIZAR TUDO" continuava a dizer "Contadores e Calendar sincronizados!", escondendo que "Calendário 2026" não tinha nenhum evento novo
- **Comando/validação:** `node --check src/Vacation_Mode.js`
- **Resultado real:** exit code 0

- **Comando/validação:** `node tests/triggers.test.js`
- **Resultado real:** 5/5 `OK`

- **Comando/validação:** `node tests/calendar.test.js`
- **Resultado real:** 9/9 `OK` (inclui o cenário novo que replica o log real: uma folha em quota, outra já sincronizada, e confirma que a notificação final identifica a folha bloqueada)

- **Comando/validação:** `node .agents/tools/validate-project.mjs`
- **Resultado real:** `{"ok":true,...}`
- **Limitações:** confirma a composição da mensagem localmente; não confirma a experiência real no Google Sheets (o toast a aparecer corretamente na folha do utilizador).

## 2026-08-20 — Quota do Calendar não detetada em conta PT / v1.5.2

- **Alteração:** `REGEX_QUOTA_CALENDAR` deixa de depender da mensagem em inglês; passa a reconhecer o sufixo `: calendar` em qualquer idioma
- **Evidência de origem:** log real de Execuções do Apps Script enviado pelo utilizador (screenshot), mostrando 7 blocos a falhar com "Serviço invocado demasiadas vezes no mesmo dia: calendar." e "Sincronização concluída (Calendário 2026): 0 criado(s), 0 atualizado(s), 0 removido(s), 7 falhado(s)."
- **Comando/validação:** `node --check src/Vacation_Mode.js`
- **Resultado real:** exit code 0

- **Comando/validação:** `node tests/triggers.test.js`
- **Resultado real:** 5/5 `OK`

- **Comando/validação:** `node tests/calendar.test.js`
- **Resultado real:** 8/8 `OK` (inclui o cenário novo com a mensagem de quota em português, replicando o log real)

- **Comando/validação:** `node .agents/tools/validate-project.mjs`
- **Resultado real:** `{"ok":true,...}`
- **Limitações:** confirma a deteção do erro localmente (Node, com o texto exato do log); não confirma ainda que a folha real cria os eventos, porque isso depende também de a quota diária real da conta Google já ter resetado — fora do controlo do código.

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
