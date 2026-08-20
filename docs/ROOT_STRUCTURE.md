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
├── src/
│   └── Vacation_Mode.js
├── tests/
│   ├── calendar.test.js
│   └── triggers.test.js
└── VERSION
```

## Política

- Ficheiros universais do produto (documentação, licença, versão) ficam na raiz.
- O código-fonte vive em `src/`; os testes locais vivem em `tests/`. A raiz fica limpa mesmo com um único ficheiro de código.
- Todo o conteúdo específico de IA vive em `.agents/` (entrada: `.agents/AGENTS.md`).
- Estado operacional vive em `.agents/state/`.
- Políticas normativas de IA vivem em `.agents/policies/`.
- A distribuição continua por cópia manual: copiar `src/Vacation_Mode.js` para o editor do Google Apps Script.
