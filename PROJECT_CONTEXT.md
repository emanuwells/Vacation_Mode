# Contexto Do Projeto

## Nome

Vacation Mode.

## Objetivo

Automatizar a gestão de férias num Google Sheets, com contadores de dias e sincronização de períodos de férias para o Google Calendar.

## Stack Técnica

| Componente | Tecnologia |
| --- | --- |
| Script principal | Google Apps Script, com sintaxe JavaScript |
| Plataforma de execução | Google Sheets e Google Calendar |
| Persistência | Próprio ficheiro Google Sheets e eventos no Google Calendar |
| Código versionado | Repositório Git local |

## Estrutura Confirmada Do Repositório

| Caminho | Função |
| --- | --- |
| `Vacation_Mode.js` | Script principal para colar no editor do Google Apps Script. |
| `README.md` | Documentação principal para utilização e manutenção. |
| `CHANGELOG.md` | Histórico versionado das alterações. |
| `CHANGELOG_POLICY.md` | Política de versionamento e changelog. |
| `AGENTS.md` | Regras gerais para IAs que trabalham no repositório. |
| `tasks/todo.md` | Plano e revisão das tarefas em curso. |
| `tasks/lessons.md` | Lições aprendidas e padrões a evitar. |
| `.gitignore` | Exclusões para metadados locais, ficheiros Office/Sheets e scripts locais. |

## Configuração Confirmada

O objeto `CONFIG`, no topo de `Vacation_Mode.js`, concentra a configuração operacional:

| Chave | Valor por omissão | Descrição |
| --- | --- | --- |
| `CALENDAR_RANGE` | `C5:AM16` | Grelha de calendário com 12 meses por 31 dias. |
| `CORES.FERIAS_ATUAL` | `#d9d2e9` | Cor de férias do ano corrente. |
| `CORES.FERIAS_ATUAL_ALT` | `#b4a7d6` | Variante aceite para férias do ano corrente. |
| `CORES.FERIAS_ANTERIOR` | `#fff2cc` | Cor para dias transitados do ano anterior. |
| `CORES.ANIVERSARIO` | `#d9ead3` | Cor para o dia de aniversário. |
| `CALENDARIO.NOME` | vazio | Usa o calendário principal quando fica vazio. |
| `CALENDARIO.TITULO_EVENTO` | `Férias` | Título-base dos eventos criados. |
| `CALENDARIO.MARCADOR` | `[FERIAS_AUTO]` | Marcador usado para identificar eventos criados pelo script. |

## Fluxos Críticos

- `onOpen()` cria o menu `Gestão de Férias` no Google Sheets.
- `sincronizarTudo()` atualiza contadores e sincroniza todos os calendários anuais encontrados.
- `atualizarContadores()` calcula férias gozadas, planeadas, totais, restantes e contadores de aniversário.
- `sincronizarComCalendar()` cria eventos no Google Calendar, agrupando datas consecutivas.
- `instalarTriggerAutomatico()` instala trigger temporal de 5 minutos e trigger `onChange`.
- `testarDetecaoCores()` ajuda a diagnosticar cores reconhecidas pelo script.

## Comandos Principais

Não há comandos locais obrigatórios confirmados. O fluxo principal é manual, através do Google Apps Script:

1. Copiar `Vacation_Mode.js` para o editor do Apps Script.
2. Guardar o projeto.
3. Recarregar o Google Sheets para disponibilizar o menu.
4. Executar as opções do menu `Gestão de Férias`.

## Política De Segredos

O repositório não deve conter credenciais, tokens, ficheiros `.clasp.json` nem `appsscript.json`. Estes ficheiros estão ignorados no `.gitignore`.

## Critérios De Verificação Antes De Concluir

- Confirmar que a documentação descreve apenas funcionalidades existentes no script.
- Confirmar que o README e este contexto não se contradizem.
- Atualizar `CHANGELOG.md` quando houver alteração versionável.
- Rever português europeu, acentuação e terminologia técnica.

## Decisões Técnicas Atuais

- O script é distribuído como ficheiro único para facilitar cópia manual para o Apps Script.
- A configuração fica centralizada no objeto `CONFIG`.
- A deteção multi-ano depende de folhas cujo nome contenha `Calendario YYYY` ou `Calendário YYYY`.
- Os eventos do Calendar são identificados pelo título configurado e pelo marcador `[FERIAS_AUTO]`.

## Riscos E Pendências

- A execução real depende de permissões do Google Apps Script e do Google Calendar.
- Não existe teste automatizado local confirmado.
- Algumas mensagens internas do script ainda usam texto sem acentuação por compatibilidade visual ou legado; qualquer alteração deve ser validada no Apps Script.

