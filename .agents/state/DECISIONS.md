# DECISIONS

Registo append-only de decisões técnicas permanentes.

## 2026-07-26 — Adotar WELLS Agent Runtime 0.5.0

- **Decisão:** concentrar todo o sistema de agentes em `.agents/`, com entrada única `.agents/AGENTS.md`.
- **Motivo:** reduzir contexto, eliminar duplicação (`AGENTS.md` raiz, `tasks/`, `docs/ai/`, `tools/ai-adapters/`) e seguir o contrato WELLS.
- **Impacto:** documentação e continuidade passam a referir `.agents/state/` e `.agents/policies/`; produção Apps Script continua por cópia manual.
- **Alternativas consideradas:** manter `AGENTS.md` na raiz com ponteiros; rejeitada por violar `ROOT_CLEAN_POLICY`.

## Modelo

- **Data:** AAAA-MM-DD
- **Decisão:**
- **Motivo:**
- **Impacto:**
- **Alternativas consideradas:**
