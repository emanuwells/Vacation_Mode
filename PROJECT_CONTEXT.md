# PROJECT_CONTEXT.md

Contexto técnico vivo do projeto.

## Identificação

| Campo | Valor |
| --- | --- |
| Nome | Vacation Mode |
| Descrição | Gestão de férias em Google Sheets com sincronização para Google Calendar |
| Versão atual | 1.4.4 |
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
| Script | Google Apps Script | Ficheiro único `Vacation_Mode.js` |
| Interface | Google Sheets | Menu `Gestão de Férias` |
| Persistência | Folha + Calendar | Sem base de dados externa |
| Distribuição | Git | Cópia manual para o Apps Script |

## Estrutura do repositório

| Caminho | Função |
| --- | --- |
| `Vacation_Mode.js` | Script principal |
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
| Sincronização ao colorir | `onAlteracaoPlanilha()` → `sincronizarTudo({ automatico: true })` |
| Contadores | `atualizarContadores()` |
| Calendar | `sincronizarComCalendar()`, `obterCalendario()`, `limparEventosAntigos()` |
| Triggers | `instalarTriggerAutomatico()`, `removerTriggerAutomatico()` |
| Diagnóstico | `testarDetecaoCores()` |

## Comandos reais

Ver `COMMANDS.md`. Validação local principal:

```bash
node --check Vacation_Mode.js
```

Validação do runtime WELLS:

```bash
node .agents/tools/validate-project.mjs
```

## Política de segredos

Não versionar credenciais, `.clasp.json` nem `appsscript.json`.

## Critérios de verificação

- Documentação alinhada com `Vacation_Mode.js`.
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

## Riscos

| Risco | Mitigação |
| --- | --- |
| Permissões Google | Autorizar script na primeira execução |
| Sem testes automatizados no runtime Google | Validação manual documentada no README |
| Triggers antigos com handler obsoleto | Reativar sincronização automática após atualizar o script |

## Pendências

- Não existe pipeline CI/CD confirmado.
- Não existe `docs/architecture/` completo; não é necessário para a escala atual do projeto.
- Produção (Apps Script) atualiza-se por cópia manual de `Vacation_Mode.js`.
