# COMMANDS.md

Comandos reais de validação e manutenção deste repositório.

## Sintaxe do script

```bash
node --check Vacation_Mode.js
```

Valida a sintaxe JavaScript do ficheiro principal. Não executa o Google Apps Script.

## Git

```bash
git status
git diff
git diff --check
```

### Commits sem cursoragent

- Nunca adicionar `Co-authored-by: Cursor <cursoragent@cursor.com>` nem commits em nome de agentes IA.
- No Cursor: **Settings → Agents → Attribution** — desativar Commit Attribution e PR Attribution.
- Com adaptador Cursor ativo, `tools/ai-adapters/cursor/.cursor/cli.json` mantém `commitAttribution` e `prAttribution` a `false`.
- Para limpar histórico existente: `scripts/strip-coauthor-msg.ps1` com `git filter-branch --msg-filter` (ver documentação Git).

## Instalação e execução

Este projeto não tem build local nem dependências npm. A instalação é manual:

1. Copiar `Vacation_Mode.js` para o editor do Google Apps Script associado à folha de cálculo.
2. Guardar o projeto no Google Apps Script.
3. Recarregar a folha e usar o menu `Gestão de Férias`.

## Notas

- Não declarar validações como executadas se estes comandos não forem corridos.
- A execução real depende de permissões do Google Sheets e do Google Calendar.
