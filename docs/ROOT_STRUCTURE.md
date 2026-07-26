# Estrutura da Raiz

Estrutura mínima deste repositório, alinhada com o contrato WELLS (`.agents/AGENTS.md`).

```text
.
├── .agents/                 # sistema IA (entrada: AGENTS.md)
├── .github/
│   └── SECURITY.md
├── CHANGELOG.md
├── COMMANDS.md
├── CONTRIBUTING.md
├── docs/
│   └── ROOT_STRUCTURE.md
├── LICENSE
├── PROJECT_CONTEXT.md
├── README.md
├── scripts/
│   └── strip-coauthor-msg.ps1
├── SECURITY.md
├── Vacation_Mode.js
└── VERSION
```

## Política

- Ficheiros universais do produto ficam na raiz.
- Todo o conteúdo específico de IA vive em `.agents/` (entrada: `.agents/AGENTS.md`).
- Estado operacional vive em `.agents/state/`.
- Políticas normativas de IA vivem em `.agents/policies/`.
- O script principal permanece na raiz enquanto a distribuição for por cópia manual.
