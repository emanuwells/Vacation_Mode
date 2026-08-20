# Changelog

## [1.5.3] - 2026-08-20T20:00:00+01:00

### "SINCRONIZAR TUDO" deixa de anunciar sucesso quando uma folha fica bloqueada por quota

**Motivo:**
- Depois da correção de idioma (1.5.2), o utilizador enviou um novo log: "Calendário 2026" ficou corretamente bloqueado por quota (mensagem e retentativa agendada, como esperado), mas "Calendário 2025" sincronizou sem erro com "0 criado(s), 0 atualizado(s), 0 removido(s), 0 falhado(s)." — porque os eventos de 2025 já existiam corretamente no Calendar de uma sincronização anterior, e a comparação por diferença não teve nada para mudar. No fim, `sincronizarTudo` mostrava sempre "Contadores e Calendar sincronizados!", independentemente de qualquer folha ter ficado bloqueada — o utilizador via essa mensagem final tranquilizadora e não percebia que "Calendário 2026" continuava sem nenhum evento novo.

**Impacto:**
- `sincronizarComCalendar` passa a devolver o resultado de cada folha (`'ok'`, `'quota'`, `'erro'` ou `'ocupado'`) em vez de nada.
- `sincronizarTudo` usa esses resultados para compor a notificação final: só diz "Contadores e Calendar sincronizados!" quando todas as folhas terminam realmente sincronizadas; caso contrário identifica pelo nome quais folhas ficaram bloqueadas por quota (com a retentativa automática já agendada) ou tiveram erro, em vez de uma mensagem genérica de sucesso.
- Não há alteração ao comportamento por folha (toasts individuais de "Sincronizar com Calendar" continuam iguais); a mudança é só na mensagem final de "SINCRONIZAR TUDO", que agora reflete com rigor o resultado real.

**Alterações:**
- `src/Vacation_Mode.js`: `sincronizarComCalendar` devolve `{ status, resultado? }` em cada ponto de saída; `sincronizarTudo` recolhe os resultados por folha; nova `construirMensagemResumoSincronizacao`.
- `tests/calendar.test.js`: novo cenário que replica o log real (uma folha bloqueada por quota, outra já sincronizada sem alterações) e confirma que a notificação final identifica a folha bloqueada, em vez de dizer "sincronizado".
- `VERSION`: `1.5.3`.

**Validação:**
- `node --check src/Vacation_Mode.js`
- `node tests/triggers.test.js` → 5/5 OK
- `node tests/calendar.test.js` → 9/9 OK (inclui o cenário novo do resumo de "SINCRONIZAR TUDO")
- `node .agents/tools/validate-project.mjs` → `ok: true`

**Diff:**
- `sincronizarComCalendar` ganha um contrato de retorno; nenhuma função existente perde comportamento. Sem alterações a menus, células ou formato de eventos.

---

## [1.5.2] - 2026-08-20T18:00:00+01:00

### Deteção de quota do Calendar independente do idioma da conta

**Motivo:**
- O utilizador enviou o log real de execução do Apps Script: a sincronização corria até ao fim ("Sincronização concluída (Calendário 2026): 0 criado(s), 0 atualizado(s), 0 removido(s), 7 falhado(s).") sem criar nenhum evento. Causa raiz confirmada pelo log: `REGEX_QUOTA_CALENDAR` só reconhecia a mensagem de quota em inglês (`Service invoked too many times for one day: calendar`), mas a conta/projeto Apps Script deste utilizador corre em português, e o Google devolve `Serviço invocado demasiadas vezes no mesmo dia: calendar.` — mensagem nunca reconhecida como quota. Cada um dos 7 blocos falhava individualmente, sem o bloqueio/retentativa automática (construídos em 1.5.0/1.5.1) alguma vez disparar, e sem o utilizador ver a notificação explicativa.

**Impacto:**
- A quota do Calendar volta a ser reconhecida corretamente, independentemente do idioma da conta Google: o Google não traduz o identificador interno do serviço (`calendar`), só a frase à volta — a nova expressão (`/:\s*calendar\b/i`) usa esse sufixo estável em vez de depender do texto em inglês.
- Quando a quota diária esgota, a sincronização volta a parar no primeiro bloco (em vez de tentar e falhar nos 7), grava o bloqueio, agenda a retentativa automática e mostra a notificação "Quota diária do Google Calendar esgotada..." em execuções manuais, em vez de um "7 falhado(s)" sem contexto.
- A quota diária real da conta Google continua sujeita ao ciclo de reposição do Google — esta correção garante que, assim que resetar, a próxima sincronização (manual ou automática) cria os eventos normalmente, sem intervenção adicional.

**Alterações:**
- `src/Vacation_Mode.js`: `REGEX_QUOTA_CALENDAR` reescrito para não depender do idioma.
- `tests/calendar.test.js`: novo cenário de regressão com a mensagem de erro em português tal como reportada pelo utilizador.
- `VERSION`: `1.5.2`.

**Validação:**
- `node --check src/Vacation_Mode.js`
- `node tests/triggers.test.js` → 5/5 OK
- `node tests/calendar.test.js` → 8/8 OK (inclui o novo cenário em português)
- `node .agents/tools/validate-project.mjs` → `ok: true`

**Diff:**
- Alteração de uma linha (expressão regular) mais o teste de regressão correspondente; nenhuma outra lógica, menu, célula ou formato de evento foi alterado.

---

## [1.5.1] - 2026-08-20T12:00:00+01:00

### Isolamento de erros por bloco na sincronização do Calendar; script movido para `src/`

**Motivo:**
- Depois de ativar a sincronização automática (1.5.0) e correr "SINCRONIZAR TUDO" na folha real, a sincronização com o Calendar continuou a falhar. Revisão do código encontrou uma regressão introduzida em 1.5.0: `sincronizarBlocosComDiferenca` já não tinha o isolamento por bloco que a versão anterior tinha (cada bloco era criado dentro do seu próprio `try/catch`); um erro num único bloco — por exemplo, um evento antigo já inválido, criado por uma versão anterior do script — abortava a sincronização inteira da folha, incluindo blocos perfeitamente válidos, e mostrava só "Erro ao sincronizar. Verifica o log.", sem indicar qual bloco falhou.
- Pedido de organização: mover `Vacation_Mode.js` para `src/`, para limpar a raiz do repositório.

**Impacto:**
- Cada bloco de férias volta a sincronizar isoladamente: um erro num bloco fica registado e contado (`falhados`), mas os restantes blocos da mesma folha continuam a ser criados/atualizados/removidos normalmente.
- Um erro de quota do Calendar continua a interromper de imediato (não faz sentido continuar a tentar quando a quota já esgotou) — é a única exceção propagada ao chamador, que trata o bloqueio e a retentativa automática.
- A notificação de "Sincronizar com Calendar" passa a incluir também os blocos falhados, quando existirem, com indicação para consultar o log.
- `src/Vacation_Mode.js` substitui `Vacation_Mode.js` na raiz; a raiz do repositório fica limpa (só documentação, licença e configuração). A instalação no Apps Script passa a copiar o conteúdo de `src/Vacation_Mode.js`.

**Alterações:**
- `Vacation_Mode.js` → `src/Vacation_Mode.js` (movido via `git mv`, sem alteração do nome do ficheiro).
- `src/Vacation_Mode.js`: `REGEX_QUOTA_CALENDAR` (constante partilhada); `sincronizarBlocosComDiferenca` volta a isolar cada mutação num `try/catch`, com `registarFalha()` a distinguir um erro de quota (propagado) de qualquer outro erro (registado e contado em `falhados`, sem abortar); mensagens de log e de notificação atualizadas para refletir blocos falhados.
- `tests/triggers.test.js`, `tests/calendar.test.js`: caminho do script atualizado para `src/Vacation_Mode.js`; novo cenário em `calendar.test.js` cobre o isolamento por bloco (um erro pontual não aborta os restantes; um erro de quota interrompe de imediato).
- `README.md`, `PROJECT_CONTEXT.md`, `COMMANDS.md`, `docs/ROOT_STRUCTURE.md`: caminhos atualizados para `src/Vacation_Mode.js`; estrutura do repositório e política da raiz refletem `src/` e `tests/`.
- `VERSION`: `1.5.1`.

**Validação:**
- `node --check src/Vacation_Mode.js`
- `node tests/triggers.test.js` → 5/5 OK
- `node tests/calendar.test.js` → 7/7 OK (inclui os 2 cenários novos de isolamento de erro por bloco)
- `node .agents/tools/validate-project.mjs` → `ok: true`

**Diff:**
- Menus, células de configuração, cores detetadas e formato dos eventos mantêm-se inalterados. A sincronização por diferença introduzida em 1.5.0 mantém-se; só a resiliência por bloco foi restaurada e o ficheiro principal mudou de localização.

---

## [1.5.0] - 2026-08-20T00:00:00+01:00

### Sincronização do Calendar por diferença, debounce ao pintar e recuperação de quota

**Motivo:**
- As datas pintadas deixavam de chegar ao Google Calendar mesmo com a sincronização automática ativa. Causa raiz: cada pintura (trigger `onChange`) e o trigger incondicional de 5 em 5 minutos disparavam uma reconstrução completa do calendário do ano inteiro (apagar todos os eventos gerados pelo script e recriá-los do zero), por cada folha anual. Isto esgotava a quota diária de invocações da Calendar API; ao esgotar, o Google lança `Service invoked too many times for one day: calendar.`, que era só registado no log (execuções automáticas correm em silêncio) — a sincronização parava sem qualquer aviso visível.
- Solução pedida pelo utilizador: aplicar o mesmo padrão já validado no projeto irmão `Luna_Sheet` (debounce, sincronização por diferença, backoff de quota com retentativa automática), preparado para folhas de anos futuros sem alterações ao código.

**Impacto:**
- A sincronização com o Calendar deixa de apagar e recriar o ano inteiro a cada execução: cria só os blocos em falta, atualiza só os que mudaram e remove só os que já não correspondem a nenhuma célula pintada.
- Pintar várias células seguidas já não dispara uma sincronização completa por pintura: os contadores continuam a atualizar de imediato, mas o Calendar sincroniza com um pequeno atraso configurável (`CONFIG.SYNC.EDIT_SYNC_DELAY_MS`, 5 min por omissão), agregando as alterações.
- O trigger periódico incondicional passa de 5 em 5 minutos para diário (rede de segurança); a frescura normal vem do debounce ao pintar.
- Se a quota diária do Calendar esgotar, o sistema grava um bloqueio em `PropertiesService` e agenda automaticamente uma retentativa (`CONFIG.SYNC.QUOTA_RETRY_DELAY_MS`, 6 h por omissão); execuções automáticas respeitam o bloqueio em silêncio, execuções manuais (`SINCRONIZAR TUDO`, `Sincronizar com Calendar`) limpam-no e forçam uma tentativa real, com um toast a explicar o motivo em vez do genérico "Erro ao sincronizar".
- O bloqueio de quota é por script (não por folha/ano), por isso uma folha nova de um ano futuro (ex. "Calendário 2027") herda automaticamente esta proteção sem exigir alterações ao código — `obterFolhasCalendario()` já a deteta pelo nome.
- Um único evento por bloco pode ficar temporariamente sem "chave" ao migrar de uma sincronização antiga (pré-1.5.0) para esta versão; a primeira sincronização depois de colar o script atualizado pode recriar esses eventos uma única vez (custo esperado de migração, não recorrente).

**Alterações:**
- `Vacation_Mode.js`: `CONFIG.SYNC` (novo), `PROPRIEDADES`, `HANDLERS`; `onAlteracaoPlanilha` separa contadores (imediatos) de Calendar (debounced); `agendarSincronizacaoCalendar`, `agendarTriggerUnico`, `removerTriggersPorHandler`, `sincronizarCalendarPendente` (novos); `sincronizarComCalendar` reescrito com verificação/gravação de quota e sincronização por diferença; `construirEventoDesejado`, `eventoGeradoPeloScript`, `obterEventosGeradosNoAno`, `sincronizarBlocosComDiferenca` (novos, substituem `limparEventosAntigos`); `instalarTriggerAutomatico` troca o trigger de 5 em 5 minutos por um diário e limpa o bloqueio de quota ao (re)instalar; `removerTriggersAutomaticos` inclui o novo handler `sincronizarCalendarPendente`; `mostrarAjuda()` atualizado.
- `tests/triggers.test.js`, `tests/calendar.test.js` (novos): cobrem debounce, cadência dos triggers, sincronização por diferença e recuperação de quota, ao estilo já usado em `Luna_Sheet/tests/`.
- `README.md`, `PROJECT_CONTEXT.md`, `COMMANDS.md`: documentação, tabela de configuração e resolução de problemas atualizadas.
- `VERSION`: `1.5.0`.

**Validação:**
- `node --check Vacation_Mode.js`
- `node tests/triggers.test.js` → 5/5 OK
- `node tests/calendar.test.js` → 5/5 OK
- `node .agents/tools/validate-project.mjs` → `ok: true`

**Diff:**
- Menus, células de configuração (`CONFIG.CELULAS`), cores detetadas e formato dos eventos criados no Calendar mantêm-se inalterados; só a estratégia interna de sincronização e a cadência dos triggers automáticos mudam.

---

## [1.4.4] - 2026-07-26T23:09:45+01:00

### Migração para WELLS Agent Runtime 0.5.0

**Motivo:**
- Unificar o contrato de agentes em `.agents/AGENTS.md` e eliminar caminhos legados (`AGENTS.md` na raiz, `tasks/`, `docs/ai/`, `tools/ai-adapters/`).

**Impacto:**
- Entrada única de IA em `.agents/AGENTS.md` (toolkit 0.5.0).
- Estado operacional em `.agents/state/`; políticas em `.agents/policies/`.
- Documentação do produto alinhada com a nova estrutura; lógica do script inalterada.

**Alterações:**
- Adicionado `.agents/` com runtime WELLS 0.5.0 (`manifest.json`, `toolkit-lock.json`).
- Removidos `AGENTS.md` (raiz), `tasks/`, `docs/ai/`, `tools/ai-adapters/`.
- Adicionados `CONTRIBUTING.md` e `SECURITY.md` na raiz.
- `README.md`, `PROJECT_CONTEXT.md`, `COMMANDS.md`, `docs/ROOT_STRUCTURE.md`: caminhos e versão atualizados.
- `Vacation_Mode.js`, `VERSION`: cabeçalho e versão `1.4.4`.

**Validação:**
- `node --check Vacation_Mode.js`
- `node .agents/tools/validate-project.mjs` → `ok: true`, versão `0.5.0`

**Diff:**
- Reorganização estrutural de agentes/docs; sem alteração funcional nos fluxos Sheets/Calendar.

---

## [1.4.3] - 2026-06-26T16:00:00+01:00

### Alinhamento com estrutura do template (raiz limpa)

**Motivo:**
- `.cursor/` e `.githooks/` na raiz violam a política de raiz limpa do template; a prevenção de `cursoragent` deve viver em `tools/ai-adapters/`.

**Impacto:**
- Raiz sem pastas de adaptador ativo; configuração Cursor opcional em `tools/ai-adapters/cursor/`.
- Regra explícita em `AGENTS.md` e `COMMANDS.md` contra co-author de agentes IA.

**Alterações:**
- Removidos `.cursor/`, `.githooks/` e `scripts/install-git-hooks.ps1` da raiz.
- `tools/ai-adapters/cursor/.cursor/cli.json`: attribution desativada no adaptador.
- `AGENTS.md`, `COMMANDS.md`, `docs/ROOT_STRUCTURE.md`: documentação atualizada.
- `VERSION`: atualizado para `1.4.3`.

**Validação:**
- Estrutura alinhada com `docs/ROOT_STRUCTURE.md`.
- Histórico Git sem trailers `cursoragent@cursor.com`.

---

## [1.4.2] - 2026-06-26T15:00:00+01:00

### Título com dias úteis e fins de semana contíguos no Calendar

**Motivo:**
- Na folha só se pintam dias úteis, mas o evento no Google Calendar deve cobrir fins de semana contíguos; o título deve refletir apenas os dias de férias contabilizados.

**Impacto:**
- O título passa a mostrar dias úteis pintados (ex.: `Férias (12 dias)`), não o total de dias de calendário.
- Eventos alargam-se a sábado/domingo adjacentes quando o período começa numa segunda ou termina numa sexta.

**Alterações:**
- `Vacation_Mode.js`: função `estenderIntervaloComFinsDeSemanaContiguos`; título e descrição com dias úteis.
- `README.md`: agrupamento documentado com contagem em dias úteis.
- `.githooks/prepare-commit-msg`, `.cursor/cli.json`: prevenção de co-author Cursor.
- `VERSION`: atualizado para `1.4.2`.

**Validação:**
- `node --check Vacation_Mode.js`
- Revisão lógica: seg–sex (5 úteis, span seg–dom); duas semanas (10 úteis, span com fins de semana).

**Diff:**
- Título deixa de usar `contarDiasCalendario`; intervalo do evento inclui fins de semana nas extremidades.

---

## [1.4.1] - 2026-06-26T14:00:00+01:00

### Agrupamento de férias com fins de semana

**Motivo:**
- Períodos de férias pintados em semanas seguidas apareciam no Google Calendar como blocos separados (ex.: dois eventos de 5 dias) porque os fins de semana não são pintados na grelha base do [Economia e Finanças](https://economiafinancas.com).

**Impacto:**
- Semanas consecutivas de férias passam a gerar um único evento com a duração total em dias de calendário (ex.: 14 dias em vez de 5+5).
- O README passa a referenciar a origem do calendário com feriados.

**Alterações:**
- `Vacation_Mode.js`: funções `intervaloApenasFinsDeSemana` e `contarDiasCalendario`; agrupamento alargado.
- `README.md`: referência ao calendário Excel do Economia e Finanças.
- `VERSION`: atualizado para `1.4.1`.

**Validação:**
- `node --check Vacation_Mode.js`
- Revisão lógica de agrupamento com cenários segunda–sexta + segunda–sexta.

**Diff:**
- Fins de semana entre blocos de dias pintados deixam de partir o período no Calendar.

---

## [1.4.0] - 2026-06-26T12:00:00+01:00

### Sincronização ao colorir e alinhamento com template

**Motivo:**
- Alinhar o repositório com a estrutura mínima do template de repositório.
- Restaurar a sincronização com o Google Calendar quando o utilizador pinta células na grelha anual.

**Impacto:**
- A documentação passa a seguir o contrato do `AGENTS.md` do template, com políticas em `docs/ai/policies/`.
- A sincronização automática deixa de entrar em loop quando o script atualiza contadores.
- Folhas com nomes como `Calendário de férias` passam a ser detetadas corretamente.
- O calendário principal é usado de imediato quando `CONFIG.CALENDARIO.NOME` está vazio.

**Alterações:**
- `AGENTS.md`: substituído pelo contrato operacional do template, com exceções documentadas para este projeto.
- `COMMANDS.md`: criado com comandos reais de validação.
- `VERSION`: criado com `1.4.0`.
- `.github/SECURITY.md`: criado.
- `docs/ROOT_STRUCTURE.md`: criado.
- `docs/ai/policies/CHANGELOG_POLICY.md`: política movida da raiz.
- `CHANGELOG_POLICY.md`: removido da raiz.
- `README.md`: reescrito de forma agnóstica e profissional.
- `PROJECT_CONTEXT.md`: reescrito com estrutura do template.
- `Vacation_Mode.js`: adicionados `onAlteracaoPlanilha`, supressão de onChange, deteção alargada de folhas e correção de `obterCalendario`.
- `tasks/todo.md` e `tasks/lessons.md`: atualizados.

**Validação:**
- `node --check Vacation_Mode.js`
- Revisão cruzada entre documentação e código.

**Diff:**
- Estrutura mínima do template + correção da sincronização ao colorir via handler dedicado de `onChange`.

---

## [1.3.5] - 2026-05-18
### Motivo
- Alinhar a documentação principal e a política de changelog com as regras atuais do `AGENTS.md`.
- Corrigir inconsistências textuais e referências obsoletas identificadas no diagnóstico do repositório.

### Alterado
- `README.md` passou a incluir badges reais no topo, incluindo a versão `1.3.5`, e documentação das funções de manutenção existentes.
- `CHANGELOG_POLICY.md` passou a exigir motivo, impacto, ficheiros alterados, testes, validação, referências e diff resumido.
- `PROJECT_CONTEXT.md` passou a registar `configurarSheet()` e `atualizarCoresAutomaticamente()` como fluxos confirmados.
- `tasks/todo.md` passou a refletir a tarefa atual de alinhamento documental.
- `Vacation_Mode.js` teve o cabeçalho atualizado para `1.3.5`, comentários técnicos acrescentados e mensagens internas corrigidas em português europeu.

### Impacto
- Melhora o onboarding, reduz contradições entre documentação e código e remove a referência textual à chave inexistente `CONFIG.CORES.FERIAS`.
- Não altera a lógica operacional de contagem, triggers ou sincronização com Google Calendar.

### Ficheiros Alterados
- `README.md`
- `CHANGELOG_POLICY.md`
- `PROJECT_CONTEXT.md`
- `tasks/todo.md`
- `Vacation_Mode.js`
- `CHANGELOG.md`

### Testes E Validação
- Revisão cruzada manual entre `AGENTS.md`, `README.md`, `PROJECT_CONTEXT.md`, `CHANGELOG_POLICY.md` e `Vacation_Mode.js`.
- Pesquisa local por referências obsoletas a `CONFIG.CORES.FERIAS` e por mensagens com caracteres corrompidos conhecidos.

### Referências
- Pedido do utilizador: avançar com as alterações pretendidas após diagnóstico.

### Diff Resumido
- Adicionados badges de stack, runtime, versão e licença, mais a secção de funções de manutenção ao README.
- Expandida a política de changelog.
- Atualizado contexto técnico e plano de trabalho.
- Corrigidas mensagens, comentários e versão no script.

## [1.3.4] - 2026-05-16
### Adicionado
- `PROJECT_CONTEXT.md` com contexto técnico confirmado do projeto, configuração operacional, fluxos críticos, política de segredos e critérios de verificação.
- `CHANGELOG_POLICY.md` com regras SemVer e formato mínimo para entradas futuras.
- `tasks/todo.md` e `tasks/lessons.md` para cumprir o fluxo de trabalho definido em `AGENTS.md`.

### Alterado
- `README.md` reescrito como documentação principal profissional, com visão geral, requisitos, instalação, configuração, utilização, menus, arquitetura, validação, troubleshooting, segurança e manutenção.

### Validação
- Conteúdo cruzado com `Vacation_Mode.js`, `AGENTS.md`, `.gitignore` e histórico existente do `CHANGELOG.md`.
- Revisão manual de coerência, acentuação e português europeu na documentação alterada.

## [1.3.3] - 2026-04-26
### Corrigido
- A sincronização automática passa a instalar também um trigger `onChange`, para apanhar alterações de formatação/cor quando se pintam dias no Google Sheets.
- A deteção de dias pintados passa a aceitar números, texto numérico e células com datas reais formatadas como dia.
- O diagnóstico de cores deixou de referir a chave inexistente `CONFIG.CORES.FERIAS`.

## [1.3.2] - 2025-12-17
### Adicionado
- Placeholders e documentação para replicação fácil, com configuração e uso consolidados no README.
- Descrições de eventos no Calendar com acentuação e formatação limpas.
- Referência ao calendário base em Excel com Feriados, de https://economiafinancas.com/.

### Alterado
- `Vacation_Mode.js` preparado para multi-ano e valores genéricos por defeito, usando o calendário principal.
- `docs/guia_rapido.md` removido; README ampliado com instruções completas.

## [1.3.1] - 2025-12-17
### Adicionado
- Suporte multi-folha e multi-ano documentado, para folhas `Calendario YYYY`.
- README reescrito e guia rápido em `docs/guia_rapido.md`, para replicação simples.

### Alterado
- Cabeçalho do script atualizado para 1.3.1.
- `Manual_Instrucoes.md` removido; informação consolidada no README.

## [1.3.0] - 2025-12-16
### Adicionado
- Genericidade: o script foi refatorado para poder ser utilizado por qualquer pessoa.
- Configuração dinâmica: o ano é agora detetado automaticamente.
- Deteção de URL: o link para o Sheet nos eventos do calendário é gerado automaticamente.
- Tratamento de erros: melhoria nas mensagens quando o calendário não é encontrado.

### Alterado
- Autor atualizado para Emanuel Ferreira (@emanuwells).
- Limpeza: remoção de emails e nomes hardcoded do código-fonte.

## [1.2.2] - 2025-11-24
### Alterado
- Correção na lógica de contagem de dias passados e futuros.
- Ajuste nas cores de deteção, para incluir variantes de roxo.

## [1.2.0] - 2025-11-01
### Adicionado
- Agrupamento de dias consecutivos no Calendar.
- Menu personalizado com opções de diagnóstico.

## [1.0.0] - 2025-01-01
### Adicionado
- Versão inicial do sistema de gestão de férias.
