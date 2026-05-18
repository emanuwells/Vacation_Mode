# Plano De Trabalho

## Tarefa Atual: Alinhar Documentação Com AGENTS.md

### Plano

- [x] Ler `AGENTS.md` e os ficheiros obrigatórios do projeto.
- [x] Identificar lacunas entre `AGENTS.md`, `README.md`, `PROJECT_CONTEXT.md`, `CHANGELOG_POLICY.md`, `CHANGELOG.md` e `Vacation_Mode.js`.
- [x] Adicionar badges reais ao topo do `README.md`.
- [x] Documentar funções operacionais existentes que estavam pouco visíveis.
- [x] Atualizar `CHANGELOG_POLICY.md` com o formato completo exigido.
- [x] Atualizar `PROJECT_CONTEXT.md` com fluxos e decisões confirmados.
- [x] Corrigir texto técnico, mensagens e referências obsoletas no script.
- [x] Atualizar `CHANGELOG.md` com entrada versionada.
- [x] Rever ortografia, acentuação e coerência entre documentação e código.
- [x] Executar verificações locais possíveis.

### Critérios De Verificação

- O `README.md` tem badges baseados apenas em stack e licença confirmadas.
- O `README.md` descreve `configurarSheet()` e `atualizarCoresAutomaticamente()` sem inventar comportamento.
- O `CHANGELOG_POLICY.md` exige motivo, impacto, ficheiros alterados, testes, validação, refs e diff resumido.
- O `PROJECT_CONTEXT.md` fica coerente com o script real.
- O script deixa de referir `CONFIG.CORES.FERIAS`, que não existe.
- A documentação e comentários ficam em português europeu correto.

### Revisão

- `README.md` recebeu badges reais e uma secção para funções de manutenção confirmadas no script.
- `PROJECT_CONTEXT.md` e `CHANGELOG_POLICY.md` foram alinhados com as regras atuais do `AGENTS.md`.
- `Vacation_Mode.js` foi atualizado para `1.3.5`, com mensagens corrigidas e sem referência operacional à chave inexistente `CONFIG.CORES.FERIAS`.
- `CHANGELOG.md` recebeu a entrada `1.3.5`, datada de 2026-05-18.
- `git diff --check` passou sem erros; restam apenas avisos normais de conversão futura LF para CRLF no Windows.
- `node --check Vacation_Mode.js` passou sem erros de sintaxe.
- Pesquisas locais confirmaram que `Vacation_Mode.js` já não contém `CONFIG.CORES.FERIAS` nem os padrões de texto corrompido pesquisados.
