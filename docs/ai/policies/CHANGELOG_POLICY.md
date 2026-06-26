# CHANGELOG_POLICY.md

Política de changelog.

## Regra principal

Qualquer alteração versionável deve ser registada em `CHANGELOG.md`.

## Alterações versionáveis

- código;
- documentação;
- configuração;
- dependências;
- Docker;
- CI/CD;
- scripts;
- estrutura;
- policies;
- remoção/renomeação de ficheiros;
- decisões técnicas.

## SemVer

- PATCH: correções e ajustes pequenos.
- MINOR: nova funcionalidade compatível.
- MAJOR: alteração estrutural/incompatível.

## Formato

```markdown
## [VERSAO] - YYYY-MM-DDTHH:mm:ss+TZ

### Título

**Motivo:**
Texto.

**Impacto:**
Texto.

**Alterações:**
- `ficheiro`: descrição.

**Validação:**
- comando ou justificação.

**Diff:**
Resumo.
---
```

## Regras

- Nunca apagar histórico antigo.
- Não inventar validações que não foram executadas.
- Manter `VERSION` alinhado com a versão mais recente do changelog.
