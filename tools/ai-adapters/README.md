# Adaptadores IA/IDE

Esta pasta guarda adaptadores opcionais para ferramentas específicas. Não devem estar ativos na raiz por defeito.

## Adaptadores

| Adaptador | Ficheiros ativados |
|---|---|
| `cursor` | `.cursor/`, `.cursorrules` |

## Ativar

Usar os scripts do [template de repositório](https://github.com/emanuwells) quando disponíveis:

```powershell
./scripts/activate-ai-adapter.ps1 -Adapter cursor
```

## Regra

O adaptador traduz o núcleo para a ferramenta. A fonte de verdade continua a ser `AGENTS.md`, `PROJECT_CONTEXT.md`, `COMMANDS.md` e `docs/`.
