# PROJECT_CONTEXT.md

Contexto técnico vivo do projeto.

## Identificação

| Campo | Valor |
| --- | --- |
| Nome | Vacation Mode |
| Descrição | Gestão de férias em Google Sheets com sincronização para Google Calendar |
| Versão atual | 1.5.1 |
| Estado | manutenção |
| Sistema IA | WELLS Agent Runtime 0.5.0 (`.agents/`) |

## Domínio

- **Problema:** contabilizar dias de férias num calendário anual e refletir períodos no Google Calendar.
- **Utilizadores:** pessoa ou equipa que gere férias numa folha de cálculo Google.
- **Regras críticas:** apenas cores configuradas contam como férias ou aniversário; eventos gerados pelo script são identificados por título e marcador `[FERIAS_AUTO]`.
- **Integrações:** Google Sheets, Google Calendar, Google Apps Script triggers.

## Stack

| Camada | Tecnologia | Observações |
| --- | --- | --- |
| Script | Google Apps Script | Ficheiro único `src/Vacation_Mode.js` |
| Interface | Google Sheets | Menu `Gestão de Férias` |
| Persistência | Folha + Calendar | Sem base de dados externa |
| Distribuição | Git | Cópia manual para o Apps Script |

## Estrutura do repositório

| Caminho | Função |
| --- | --- |
| `src/Vacation_Mode.js` | Script principal |
| `tests/` | Testes locais Node.js (triggers, sincronização por diferença, quota) |
| `README.md` | Documentação de utilização |
| `COMMANDS.md` | Comandos reais de validação |
| `CHANGELOG.md` | Histórico versionado |
| `PROJECT_CONTEXT.md` | Este ficheiro |
| `CONTRIBUTING.md` | Fluxo de contribuição |
| `SECURITY.md` | Política de segurança na raiz |
| `.github/SECURITY.md` | Política de segurança no GitHub |
| `.agents/AGENTS.md` | Contrato operacional para IAs (entrada única) |
| `.agents/state/` | TODO, HANDOFF, LESSONS, DECISIONS, EVIDENCE |
| `.agents/policies/` | Políticas normativas (incl. changelog) |
| `docs/ROOT_STRUCTURE.md` | Estrutura da raiz do projeto |
| `scripts/` | Utilitários Git locais |

## Fluxos críticos

| Fluxo | Funções |
| --- | --- |
| Abertura | `onOpen()` cria o menu |
| Sincronização manual | `sincronizarTudo()` |
| Sincronização ao colorir | `onAlteracaoPlanilha()` → contadores imediatos + `agendarSincronizacaoCalendar()` (debounce) → `sincronizarCalendarPendente()` |
| Contadores | `atualizarContadores()` |
| Calendar | `sincronizarComCalendar()`, `obterCalendario()`, `obterEventosGeradosNoAno()`, `sincronizarBlocosComDiferenca()` |
| Quota do Calendar | Deteção do erro `Service invoked too many times: calendar` em `sincronizarComCalendar()`, retentativa via `PropertiesService` + `agendarTriggerUnico()` |
| Triggers | `instalarTriggerAutomatico()` (diário + onChange), `removerTriggerAutomatico()` |
| Diagnóstico | `testarDetecaoCores()` |

## Comandos reais

Ver `COMMANDS.md`. Validação local principal:

```bash
node --check src/Vacation_Mode.js
node tests/triggers.test.js
node tests/calendar.test.js
```

Validação do runtime WELLS:

```bash
node .agents/tools/validate-project.mjs
```

## Política de segredos

Não versionar credenciais, `.clasp.json` nem `appsscript.json`.

## Critérios de verificação

- Documentação alinhada com `src/Vacation_Mode.js`.
- `CHANGELOG.md` e `VERSION` atualizados em alterações versionáveis.
- Português europeu com acentuação correta.
- Entrada de agentes apenas via `.agents/AGENTS.md`.

## Decisões técnicas

- Distribuição como ficheiro único para facilitar cópia manual.
- Configuração centralizada em `CONFIG`.
- `onChange` usa handler dedicado `onAlteracaoPlanilha` para evitar loops quando o script atualiza contadores.
- `PropertiesService` suprime reentrância durante escrita do script.
- Deteção de folhas alargada para nomes como `Calendário de férias`.
- Sistema de agentes migrado para WELLS 0.5.0 em `.agents/`.
- Sincronização do Calendar por diferença (`sincronizarBlocosComDiferenca`), em vez de apagar e recriar o ano inteiro a cada execução: reduz drasticamente as chamadas à Calendar API e foi a causa raiz de a sincronização automática parecer parar de funcionar ao esgotar a quota diária. Padrão portado do projeto irmão `Luna_Sheet`.
- Debounce da sincronização com o Calendar ao pintar (`agendarSincronizacaoCalendar` + `agendarTriggerUnico`, com `LockService`): contadores continuam imediatos; o Calendar agrega várias pinturas seguidas num único ciclo, adiado por `CONFIG.SYNC.EDIT_SYNC_DELAY_MS`.
- Trigger periódico incondicional passou de 5 em 5 minutos para diário: a frescura normal vem do debounce ao pintar; o trigger diário é só rede de segurança. O polling de 5 em 5 minutos era o maior consumidor de quota quando nada mudava.
- Backoff de quota do Calendar script-wide via `PropertiesService.getScriptProperties()` (não por documento nem por folha/ano): a quota é da execução do script, por isso uma folha nova (ex. "Calendário 2027") herda automaticamente a mesma proteção sem alterações ao código.

## Riscos

| Risco | Mitigação |
| --- | --- |
| Permissões Google | Autorizar script na primeira execução |
| Quota diária do Google Calendar esgotada | Sincronização por diferença (menos chamadas à API), debounce ao pintar, trigger diário em vez de 5 em 5 min, e retentativa automática via `PropertiesService`; execução manual força e desbloqueia |
| Sem testes automatizados no runtime Google real | `tests/triggers.test.js` e `tests/calendar.test.js` cobrem debounce, sincronização por diferença e recuperação de quota localmente (Node.js, sem depender do Apps Script); a validação no Calendar/Sheets reais continua manual |
| Triggers antigos com handler obsoleto | Reativar sincronização automática após atualizar o script |

## Pendências

- Não existe pipeline CI/CD confirmado.
- Não existe `docs/architecture/` completo; não é necessário para a escala atual do projeto.
- Produção (Apps Script) atualiza-se por cópia manual de `src/Vacation_Mode.js`.
