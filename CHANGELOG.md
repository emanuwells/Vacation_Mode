# Changelog

## [1.4.1] - 2026-06-26T14:00:00+01:00

### Agrupamento de férias com fins de semana

**Motivo:**
- Períodos de férias pintados em semanas seguidas apareciam no Google Calendar como blocos separados (ex.: dois eventos de 5 dias) porque os fins de semana não são pintados na grelha base do [Economia e Finanças](https://economiafinancas.com).

**Impacto:**
- Semanas consecutivas de férias passam a gerar um único evento com a duração total em dias de calendário (ex.: 14 dias em vez de 5+5).
- O README passa a referenciar a origem do calendário com feriados.

**Alterações:**
- `Vacation_Mode.js`: funções `intervaloApenasFinsDeSemana` e `contarDiasCalendario`; agrupamento alargado.
- `README.md`: referência ao calendário Excel do Economia e Finanças.
- `VERSION`: atualizado para `1.4.1`.

**Validação:**
- `node --check Vacation_Mode.js`
- Revisão lógica de agrupamento com cenários segunda–sexta + segunda–sexta.

**Diff:**
- Fins de semana entre blocos de dias pintados deixam de partir o período no Calendar.

---

## [1.4.0] - 2026-06-26T12:00:00+01:00

### Sincronização ao colorir e alinhamento com template

**Motivo:**
- Alinhar o repositório com a estrutura mínima do template de repositório.
- Restaurar a sincronização com o Google Calendar quando o utilizador pinta células na grelha anual.

**Impacto:**
- A documentação passa a seguir o contrato do `AGENTS.md` do template, com políticas em `docs/ai/policies/`.
- A sincronização automática deixa de entrar em loop quando o script atualiza contadores.
- Folhas com nomes como `Calendário de férias` passam a ser detetadas corretamente.
- O calendário principal é usado de imediato quando `CONFIG.CALENDARIO.NOME` está vazio.

**Alterações:**
- `AGENTS.md`: substituído pelo contrato operacional do template, com exceções documentadas para este projeto.
- `COMMANDS.md`: criado com comandos reais de validação.
- `VERSION`: criado com `1.4.0`.
- `.github/SECURITY.md`: criado.
- `docs/ROOT_STRUCTURE.md`: criado.
- `docs/ai/policies/CHANGELOG_POLICY.md`: política movida da raiz.
- `CHANGELOG_POLICY.md`: removido da raiz.
- `README.md`: reescrito de forma agnóstica e profissional.
- `PROJECT_CONTEXT.md`: reescrito com estrutura do template.
- `Vacation_Mode.js`: adicionados `onAlteracaoPlanilha`, supressão de onChange, deteção alargada de folhas e correção de `obterCalendario`.
- `tasks/todo.md` e `tasks/lessons.md`: atualizados.

**Validação:**
- `node --check Vacation_Mode.js`
- Revisão cruzada entre documentação e código.

**Diff:**
- Estrutura mínima do template + correção da sincronização ao colorir via handler dedicado de `onChange`.

---

## [1.3.5] - 2026-05-18
### Motivo
- Alinhar a documentação principal e a política de changelog com as regras atuais do `AGENTS.md`.
- Corrigir inconsistências textuais e referências obsoletas identificadas no diagnóstico do repositório.

### Alterado
- `README.md` passou a incluir badges reais no topo, incluindo a versão `1.3.5`, e documentação das funções de manutenção existentes.
- `CHANGELOG_POLICY.md` passou a exigir motivo, impacto, ficheiros alterados, testes, validação, referências e diff resumido.
- `PROJECT_CONTEXT.md` passou a registar `configurarSheet()` e `atualizarCoresAutomaticamente()` como fluxos confirmados.
- `tasks/todo.md` passou a refletir a tarefa atual de alinhamento documental.
- `Vacation_Mode.js` teve o cabeçalho atualizado para `1.3.5`, comentários técnicos acrescentados e mensagens internas corrigidas em português europeu.

### Impacto
- Melhora o onboarding, reduz contradições entre documentação e código e remove a referência textual à chave inexistente `CONFIG.CORES.FERIAS`.
- Não altera a lógica operacional de contagem, triggers ou sincronização com Google Calendar.

### Ficheiros Alterados
- `README.md`
- `CHANGELOG_POLICY.md`
- `PROJECT_CONTEXT.md`
- `tasks/todo.md`
- `Vacation_Mode.js`
- `CHANGELOG.md`

### Testes E Validação
- Revisão cruzada manual entre `AGENTS.md`, `README.md`, `PROJECT_CONTEXT.md`, `CHANGELOG_POLICY.md` e `Vacation_Mode.js`.
- Pesquisa local por referências obsoletas a `CONFIG.CORES.FERIAS` e por mensagens com caracteres corrompidos conhecidos.

### Referências
- Pedido do utilizador: avançar com as alterações pretendidas após diagnóstico.

### Diff Resumido
- Adicionados badges de stack, runtime, versão e licença, mais a secção de funções de manutenção ao README.
- Expandida a política de changelog.
- Atualizado contexto técnico e plano de trabalho.
- Corrigidas mensagens, comentários e versão no script.

## [1.3.4] - 2026-05-16
### Adicionado
- `PROJECT_CONTEXT.md` com contexto técnico confirmado do projeto, configuração operacional, fluxos críticos, política de segredos e critérios de verificação.
- `CHANGELOG_POLICY.md` com regras SemVer e formato mínimo para entradas futuras.
- `tasks/todo.md` e `tasks/lessons.md` para cumprir o fluxo de trabalho definido em `AGENTS.md`.

### Alterado
- `README.md` reescrito como documentação principal profissional, com visão geral, requisitos, instalação, configuração, utilização, menus, arquitetura, validação, troubleshooting, segurança e manutenção.

### Validação
- Conteúdo cruzado com `Vacation_Mode.js`, `AGENTS.md`, `.gitignore` e histórico existente do `CHANGELOG.md`.
- Revisão manual de coerência, acentuação e português europeu na documentação alterada.

## [1.3.3] - 2026-04-26
### Corrigido
- A sincronização automática passa a instalar também um trigger `onChange`, para apanhar alterações de formatação/cor quando se pintam dias no Google Sheets.
- A deteção de dias pintados passa a aceitar números, texto numérico e células com datas reais formatadas como dia.
- O diagnóstico de cores deixou de referir a chave inexistente `CONFIG.CORES.FERIAS`.

## [1.3.2] - 2025-12-17
### Adicionado
- Placeholders e documentação para replicação fácil, com configuração e uso consolidados no README.
- Descrições de eventos no Calendar com acentuação e formatação limpas.
- Referência ao calendário base em Excel com Feriados, de https://economiafinancas.com/.

### Alterado
- `Vacation_Mode.js` preparado para multi-ano e valores genéricos por defeito, usando o calendário principal.
- `docs/guia_rapido.md` removido; README ampliado com instruções completas.

## [1.3.1] - 2025-12-17
### Adicionado
- Suporte multi-folha e multi-ano documentado, para folhas `Calendario YYYY`.
- README reescrito e guia rápido em `docs/guia_rapido.md`, para replicação simples.

### Alterado
- Cabeçalho do script atualizado para 1.3.1.
- `Manual_Instrucoes.md` removido; informação consolidada no README.

## [1.3.0] - 2025-12-16
### Adicionado
- Genericidade: o script foi refatorado para poder ser utilizado por qualquer pessoa.
- Configuração dinâmica: o ano é agora detetado automaticamente.
- Deteção de URL: o link para o Sheet nos eventos do calendário é gerado automaticamente.
- Tratamento de erros: melhoria nas mensagens quando o calendário não é encontrado.

### Alterado
- Autor atualizado para Emanuel Ferreira (@emanuwells).
- Limpeza: remoção de emails e nomes hardcoded do código-fonte.

## [1.2.2] - 2025-11-24
### Alterado
- Correção na lógica de contagem de dias passados e futuros.
- Ajuste nas cores de deteção, para incluir variantes de roxo.

## [1.2.0] - 2025-11-01
### Adicionado
- Agrupamento de dias consecutivos no Calendar.
- Menu personalizado com opções de diagnóstico.

## [1.0.0] - 2025-01-01
### Adicionado
- Versão inicial do sistema de gestão de férias.
