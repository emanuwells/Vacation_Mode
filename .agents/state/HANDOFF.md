# HANDOFF

- **Última atualização:** 2026-08-20T18:00:00+01:00
- **Estado:** concluído — v1.5.2 em `master`, ainda por colar na folha real
- **Branch:** `master`
- **Objetivo:** corrigir a deteção de quota do Calendar, que nunca disparava numa conta em português

## Estado útil

- **Concluído:** causa raiz confirmada pelo log real de Execuções enviado pelo utilizador: `REGEX_QUOTA_CALENDAR` só reconhecia a mensagem de quota em inglês; a conta corre em português (`Serviço invocado demasiadas vezes no mesmo dia: calendar.`), por isso o bloqueio/retentativa (1.5.0/1.5.1) nunca disparava e todos os blocos falhavam individualmente sem explicação. Regex reescrito para depender só do sufixo técnico não traduzido (`: calendar`), válido em qualquer idioma. Teste de regressão com a mensagem exata em português.
- **Em curso:** N/A
- **Bloqueios/riscos:** a quota diária real da conta Google estava mesmo esgotada no momento do log (confirmado pela mensagem do próprio Google); isto é do lado do Google, não é corrigível por código. A correção garante que, quando a quota resetar, a sincronização (manual ou pela retentativa automática já agendada) finalmente cria os eventos, e que entretanto o utilizador vê a mensagem correta em vez de "N falhado(s)" sem contexto.
- **Ficheiros relevantes:** `src/Vacation_Mode.js` (`REGEX_QUOTA_CALENDAR`), `tests/calendar.test.js`
- **Validações executadas:** `node --check src/Vacation_Mode.js`; `node tests/triggers.test.js` (5/5); `node tests/calendar.test.js` (8/8, incluindo o cenário novo em português); `node .agents/tools/validate-project.mjs` (`ok: true`)
- **Próximo passo exato:** colar `src/Vacation_Mode.js` no Apps Script (substitui tudo), reativar "Ativar Sincronização Automática" e correr "SINCRONIZAR TUDO" ou "Sincronizar com Calendar". Se a quota ainda estiver esgotada nesse momento, a notificação vai dizer isso explicitamente e o sistema tenta sozinho mais tarde — não é preciso repetir manualmente. Se aparecer QUALQUER outra mensagem de erro (não relacionada com quota), reportar o texto exato para diagnóstico dirigido.
