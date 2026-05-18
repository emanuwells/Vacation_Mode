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
### Motivo
- Porque a alteração foi necessária.

### Alterado
- Descrição objetiva da alteração.

### Impacto
- Efeito esperado para utilizadores, manutenção, configuração ou execução.

### Ficheiros Alterados
- `caminho/do/ficheiro`

### Testes E Validação
- Como foi verificada a alteração.

### Referências
- Pedido, issue, PR ou contexto relevante. Usar `N/A` quando não existir referência externa.

### Diff Resumido
- Resumo curto do que mudou, sem substituir a leitura do diff.
```

Se alguma secção não tiver referência externa ou validação executável aplicável, deve ficar explícito com `N/A` ou com a indicação concreta da limitação.

## Regras

- Nunca apagar histórico antigo.
- Não inventar validações que não foram executadas.
- Registar motivo, impacto, ficheiros alterados, testes, validação, referências e diff resumido quando a alteração tocar documentação, configuração ou código.
- Se não houver alteração versionável, declarar isso na resposta final.
