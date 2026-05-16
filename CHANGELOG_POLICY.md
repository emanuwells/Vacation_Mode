# Política De Changelog

Este projeto usa versionamento SemVer (`MAJOR.MINOR.PATCH`) e mantém o histórico em `CHANGELOG.md`.

## Quando Atualizar

Atualizar `CHANGELOG.md` sempre que houver alteração versionável, incluindo:

- alterações de comportamento no `Vacation_Mode.js`;
- mudanças de configuração, instalação, fluxos ou documentação principal;
- correções de bugs;
- melhorias de documentação que afetem onboarding, operação ou manutenção.

## Como Versionar

| Tipo | Critério |
| --- | --- |
| `MAJOR` | Alterações incompatíveis ou que exigem migração manual relevante. |
| `MINOR` | Novas funcionalidades compatíveis com versões anteriores. |
| `PATCH` | Correções, ajustes documentais e melhorias compatíveis. |

## Formato Das Entradas

As entradas novas devem ficar no topo do `CHANGELOG.md` e seguir esta estrutura sempre que aplicável:

```markdown
## [X.Y.Z] - AAAA-MM-DD
### Alterado
- Descrição objetiva da alteração.

### Validação
- Como foi verificada a alteração.
```

## Regras

- Nunca apagar histórico antigo.
- Não inventar validações que não foram executadas.
- Registar ficheiros relevantes quando a alteração tocar documentação, configuração ou código.
- Se não houver alteração versionável, declarar isso na resposta final.

