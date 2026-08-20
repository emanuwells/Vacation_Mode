# Lições Aprendidas

Este ficheiro regista padrões a manter ou evitar em trabalho futuro neste repositório.

## Registos

- Nunca detetar um erro do Apps Script/Google pelo texto traduzido da mensagem (ex.: "Service invoked too many times..."); o Google traduz a frase consoante o idioma da conta, mas mantém o identificador interno do serviço em inglês (ex.: sempre termina em ": calendar"). Detetar pelo sufixo técnico estável, não pela frase completa — esta app já teve o bug (1.5.0/1.5.1 só reconheciam a mensagem em inglês; a conta do utilizador corre em português e nunca disparava o bloqueio de quota).
- Ao substituir uma estratégia "tudo ou nada" (apagar tudo/recriar tudo) por uma sincronização granular por item, preservar explicitamente o isolamento de erros por item (`try/catch` por mutação individual); sem isso, um único item problemático volta a abortar a operação inteira, anulando o ganho de granularidade. Aconteceu entre 1.5.0 (perdeu o isolamento por bloco que a versão anterior tinha) e 1.5.1 (restaurado).
- Antes de alterar documentação, confirmar sempre as funcionalidades no `src/Vacation_Mode.js` e não depender apenas do README existente.
- Quando ficheiros obrigatórios definidos em `.agents/AGENTS.md` não existirem, criá-los com factos confirmados e marcar como `A confirmar` qualquer informação não validada.
- Triggers `onChange` ligados diretamente a funções que escrevem na folha criam loops; usar handler dedicado e supressão temporária com `PropertiesService`.
- Ao atualizar handlers de triggers no script, o utilizador deve reativar a sincronização automática para substituir triggers antigos.
- Nunca versionar `.cursor/` ou `.githooks/` na raiz; adaptadores Cursor vivem em `.agents/adapters/` e só são lidos quando necessário.
- O trailer `Co-authored-by: Cursor <cursoragent@cursor.com>` faz o GitHub listar `cursoragent` como contributor; desativar Attribution no Cursor e não adicionar co-authors de agentes IA.
- Após migrar para WELLS, não recriar `AGENTS.md`, `tasks/` nem `docs/ai/` na raiz; a entrada canónica é `.agents/AGENTS.md`.
