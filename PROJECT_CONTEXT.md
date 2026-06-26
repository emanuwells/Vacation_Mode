# PROJECT_CONTEXT.md

Contexto técnico vivo do projeto.

## Identificação

| Campo | Valor |
| --- | --- |
| Nome | Vacation Mode |
| Descrição | Gestão de férias em Google Sheets com sincronização para Google Calendar |
| Versão atual | 1.4.0 |
| Estado | manutenção |

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
| `AGENTS.md` | Contrato operacional para IAs |
| `COMMANDS.md` | Comandos reais de validação |
| `CHANGELOG.md` | Histórico versionado |
| `docs/ai/policies/CHANGELOG_POLICY.md` | Política de changelog |
| `PROJECT_CONTEXT.md` | Este ficheiro |
| `tasks/` | Plano e lições aprendidas |
| `.github/SECURITY.md` | Política de segurança |

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

## Política de segredos

Não versionar credenciais, `.clasp.json` nem `appsscript.json`.

## Critérios de verificação

- Documentação alinhada com `Vacation_Mode.js`.
- `CHANGELOG.md` e `VERSION` atualizados em alterações versionáveis.
- Português europeu com acentuação correta.

## Decisões técnicas

- Distribuição como ficheiro único para facilitar cópia manual.
- Configuração centralizada em `CONFIG`.
- `onChange` usa handler dedicado `onAlteracaoPlanilha` para evitar loops quando o script atualiza contadores.
- `PropertiesService` suprime reentrância durante escrita do script.
- Deteção de folhas alargada para nomes como `Calendário de férias`.

## Riscos

| Risco | Mitigação |
| --- | --- |
| Permissões Google | Autorizar script na primeira execução |
| Sem testes automatizados no runtime Google | Validação manual documentada no README |
| Triggers antigos com handler obsoleto | Reativar sincronização automática após atualizar o script |

## Pendências

- Não existe pipeline CI/CD confirmado.
- Não existe `docs/architecture/` completo; não é necessário para a escala atual do projeto.
