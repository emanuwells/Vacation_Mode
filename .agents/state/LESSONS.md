# Lições Aprendidas

Este ficheiro regista padrões a manter ou evitar em trabalho futuro neste repositório.

## Registos

- Antes de alterar documentação, confirmar sempre as funcionalidades no `Vacation_Mode.js` e não depender apenas do README existente.
- Quando ficheiros obrigatórios definidos em `.agents/AGENTS.md` não existirem, criá-los com factos confirmados e marcar como `A confirmar` qualquer informação não validada.
- Triggers `onChange` ligados diretamente a funções que escrevem na folha criam loops; usar handler dedicado e supressão temporária com `PropertiesService`.
- Ao atualizar handlers de triggers no script, o utilizador deve reativar a sincronização automática para substituir triggers antigos.
- Nunca versionar `.cursor/` ou `.githooks/` na raiz; adaptadores Cursor vivem em `.agents/adapters/` e só são lidos quando necessário.
- O trailer `Co-authored-by: Cursor <cursoragent@cursor.com>` faz o GitHub listar `cursoragent` como contributor; desativar Attribution no Cursor e não adicionar co-authors de agentes IA.
- Após migrar para WELLS, não recriar `AGENTS.md`, `tasks/` nem `docs/ai/` na raiz; a entrada canónica é `.agents/AGENTS.md`.
