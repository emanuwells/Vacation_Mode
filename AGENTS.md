# AGENTS.md

Este ficheiro define regras gerais para IAs que trabalhem neste repositório.

Para regras específicas do projeto atual, ler também `PROJECT_CONTEXT.md`.

## Ordem De Leitura Obrigatória

1. `AGENTS.md` — regras gerais de trabalho.
2. `PROJECT_CONTEXT.md` — contexto específico do projeto.
3. `CHANGELOG_POLICY.md` — política obrigatória de versionamento e atualização automática do changelog.
4. `CHANGELOG.md` — histórico versionado das alterações.
5. `tasks/lessons.md` — lições aprendidas e erros a evitar.
6. `tasks/todo.md` — plano atual e estado da execução.
7. `README.md` — documentação principal para humanos.
8. Documentação técnica do projeto, quando existir.

## Ordem De Prioridade Em Caso De Conflito

1. Instruções explícitas do utilizador.
2. Regras específicas em `PROJECT_CONTEXT.md`.
3. Regras gerais deste `AGENTS.md`.
4. Convenções reais do código existente.
5. Preferências inferidas pela IA.

## Camada Específica De Projeto

Cada projeto deve ter um ficheiro `PROJECT_CONTEXT.md` na raiz.

Esse ficheiro deve conter:

- nome e objetivo do projeto;
- stack técnica;
- estrutura real do repositório;
- comandos principais;
- regras específicas;
- endpoints, jobs, scripts ou fluxos críticos;
- política de segredos;
- critérios de verificação antes de concluir trabalho;
- decisões técnicas atuais;
- riscos, pendências e dívida técnica.

Se o ficheiro não existir, a IA deve criá-lo antes de iniciar alterações não triviais.

Se existir mas estiver desatualizado, a IA deve atualizá-lo na mesma tarefa em que alterar o projeto.

## Regras Para Criar Ou Atualizar `PROJECT_CONTEXT.md`

- Não inventar informação.
- Separar factos confirmados de inferências.
- Marcar como `A confirmar` tudo o que não esteja validado.
- Preferir comandos verificados no repositório.
- Registar decisões técnicas que afetem trabalho futuro.
- Manter o texto em português europeu com acentuação.
- Garantir que toda a documentação e comentários técnicos ficam em português europeu correto, com acentuação.
- Atualizar documentação quando a arquitetura, comandos, variáveis de ambiente ou fluxos mudarem.

## Política De Documentação E Código Comentado

Toda a documentação do projeto deve estar escrita em português europeu, com acentuação correta e sem erros ortográficos.

Isto aplica-se a:

- `README.md`;
- `PROJECT_CONTEXT.md`;
- `CHANGELOG.md`;
- `CHANGELOG_POLICY.md`;
- ficheiros em `docs/`;
- comentários no código;
- docstrings;
- PHPDoc, JSDoc, TSDoc ou equivalente;
- mensagens explicativas em scripts internos;
- documentação de endpoints, comandos, variáveis de ambiente e fluxos técnicos.

## Regras Obrigatórias De Documentação

- Documentar o que cada módulo, script, função, endpoint ou fluxo relevante faz.
- Usar português europeu claro, técnico e sem mistura com português do Brasil.
- Usar acentuação correta em todo o texto.
- Não deixar comentários vagos como `fix`, `todo`, `hack`, `stuff`, `coisas`, sem explicação concreta.
- Explicar o motivo de decisões técnicas quando afetarem manutenção futura.
- Atualizar a documentação sempre que o comportamento do código mudar.
- Remover ou corrigir documentação obsoleta na mesma alteração em que o código for modificado.
- Preferir frases simples, objetivas e tecnicamente corretas.
- Não documentar em excesso código óbvio, mas documentar sempre regras de negócio, integrações, decisões e efeitos laterais.

## Regras Para Comentários No Código

Os comentários no código devem explicar intenção, regras de negócio, decisões técnicas ou riscos.

Comentários aceitáveis:

```php
// Valida se o agente pode enviar métricas sem autenticação na V1.
```

```ts
// Mantém compatibilidade temporária com o alias antigo `/login`.
```

Comentários a evitar:

```php
// Faz coisas
```

```ts
// Fix bug
```

```js
// TODO
```

Se existir um `TODO`, deve ter contexto mínimo:

```ts
// TODO: remover este alias quando todos os clientes usarem `/auth/login`.
```

## Qualidade Linguística Obrigatória

Antes de concluir qualquer tarefa que altere documentação ou comentários, a IA deve rever:

- ortografia;
- acentuação;
- termos técnicos;
- clareza;
- coerência com o restante projeto;
- consistência entre `README.md`, `PROJECT_CONTEXT.md`, `CHANGELOG.md` e o código.

Não entregar trabalho como concluído se a documentação ou comentários estiverem em português incorreto, sem acentos ou inconsistentes com o código.

## README Obrigatório

Cada projeto deve ter um `README.md` profissional, claro, completo e visualmente apelativo.

O `README.md` é a documentação principal para humanos. Deve permitir que uma pessoa ou equipa compreenda rapidamente o objetivo do projeto, como o executar, como o configurar, como o testar, como o manter e como contribuir sem depender de explicações externas.

Se o `README.md` não existir, a IA deve criá-lo antes de concluir qualquer alteração não trivial.

Se o `README.md` existir mas estiver incompleto, desatualizado, confuso, pouco profissional ou visualmente pobre, a IA deve melhorá-lo na mesma tarefa em que alterar o projeto.

## Qualidade Obrigatória Do `README.md`

O `README.md` deve seguir padrões profissionais equivalentes aos melhores repositórios públicos e internos.

Deve ser:

- claro;
- bem estruturado;
- visualmente limpo;
- tecnicamente rigoroso;
- fácil de percorrer;
- escrito em português europeu correto, com acentuação;
- atualizado com o estado real do projeto;
- útil para instalação, execução, manutenção e onboarding.

## Estrutura Recomendada Do `README.md`

Sempre que aplicável, o `README.md` deve incluir:

- título do projeto;
- badges visuais úteis no topo, quando houver stack, licença, versão ou estado confirmados;
- descrição curta e objetiva;
- índice, se o documento for longo;
- visão geral;
- funcionalidades principais;
- stack tecnológica;
- requisitos;
- instalação;
- configuração;
- variáveis de ambiente;
- comandos principais;
- utilização com exemplos;
- estrutura do projeto;
- arquitetura técnica;
- fluxos importantes;
- testes;
- qualidade, linting ou formatação;
- troubleshooting;
- segurança e gestão de segredos;
- roadmap ou pendências relevantes;
- contribuição, quando aplicável;
- licença;
- referência ao `CHANGELOG.md`.

## Regras Para Criar Ou Atualizar O `README.md`

- Não inventar funcionalidades, comandos, endpoints, dependências ou decisões técnicas.
- Validar comandos no repositório sempre que possível.
- Separar factos confirmados de informação marcada como `A confirmar`.
- Refletir a estrutura real do projeto.
- Manter exemplos simples e executáveis.
- Atualizar instruções quando forem alterados scripts, Docker, variáveis de ambiente, migrations, endpoints ou fluxos principais.
- Remover instruções obsoletas.
- Usar tabelas quando melhorarem a leitura.
- Usar blocos de código com linguagem indicada.
- Evitar texto genérico sem utilidade prática.
- Garantir que o `README.md` não contradiz `PROJECT_CONTEXT.md`, `CHANGELOG.md`, `CHANGELOG_POLICY.md` ou o código.

## Apresentação Visual Do `README.md`

O `README.md` deve ser visualmente apelativo sem sacrificar rigor técnico.

Boas práticas:

- usar títulos e subtítulos claros;
- usar listas curtas e objetivas;
- usar tabelas para comandos, variáveis de ambiente e endpoints;
- usar blocos de código formatados;
- incluir diagramas Mermaid quando ajudarem a explicar arquitetura ou fluxos;
- destacar avisos importantes com secções como `Nota`, `Atenção` ou `Importante`;
- evitar parágrafos longos;
- manter consistência de termos e formatação.

## Badges Obrigatórios No `README.md`

Sempre que fizer sentido, o `README.md` deve incluir badges visuais no topo do ficheiro, logo abaixo do título e antes da descrição principal.

Os badges devem resumir rapidamente:

- stack principal do projeto;
- framework ou runtime;
- base de dados;
- estado do projeto, se aplicável;
- licença;
- versão, se aplicável;
- estado dos testes ou CI/CD, se existir.

Exemplo esperado:

```md
![Stack](https://img.shields.io/badge/stack-React%20%7C%20Vite%20%7C%20PHP%20%7C%20Postgres-29b6f6)
![License](https://img.shields.io/badge/license-MIT-2ecc71)
```

Exemplo no topo de um `README.md`:

```md
# WELLS_OS

![Stack](https://img.shields.io/badge/stack-React%20%7C%20Vite%20%7C%20PHP%20%7C%20Postgres-29b6f6)
![License](https://img.shields.io/badge/license-MIT-2ecc71)

Portfolio e painel operacional do Emanuel Wells.
```

Os badges devem ser criados com base na stack real do projeto.

A IA não deve inventar tecnologias, licença, versão, estado de testes, cobertura, build ou CI/CD.

Se a licença ainda não estiver definida, usar:

```md
![License](https://img.shields.io/badge/license-A%20confirmar-lightgrey)
```

Se o projeto ainda estiver em desenvolvimento inicial, pode ser usado:

```md
![Status](https://img.shields.io/badge/status-em%20desenvolvimento-f39c12)
```

## Regras Para Badges No `README.md`

- Usar badges apenas quando acrescentarem leitura rápida ao `README.md`.
- Colocar os badges no topo do documento.
- Manter nomes curtos, claros e legíveis.
- Usar apenas tecnologias confirmadas no projeto.
- Não usar badges falsos de build, testes, cobertura, versão ou CI/CD sem validação real.
- Preferir badges simples, limpos e consistentes.
- Garantir que os badges não substituem a documentação textual.
- Atualizar os badges quando a stack, licença, estado ou versão do projeto mudarem.

## Dockerização Quando Fizer Sentido

Sempre que fizer sentido para o projeto, a IA deve propor ou implementar uma configuração Docker simples, segura e profissional.

Docker deve ser considerado especialmente quando o projeto tiver:

- backend ou API;
- base de dados;
- workers, filas ou serviços auxiliares;
- dependências difíceis de instalar manualmente;
- necessidade de ambiente reprodutível;
- deploy em VPS, Coolify, Portainer, CI/CD ou ambiente semelhante;
- diferença relevante entre ambiente local e produção.

A Dockerização deve ser simples e adequada à escala do projeto.

Não criar Docker, Docker Compose ou infraestrutura adicional se isso apenas aumentar complexidade sem benefício claro.

## Regras Para Docker

Quando Docker for usado, devem existir, sempre que aplicável:

- `Dockerfile` claro e minimalista;
- `docker-compose.yml` para ambiente local;
- `.dockerignore`;
- `.env.example`;
- instruções no `README.md`;
- documentação das portas, volumes e variáveis de ambiente;
- comandos de arranque, paragem, rebuild e logs;
- separação clara entre ambiente local, desenvolvimento e produção quando aplicável.

A IA não deve colocar segredos reais em imagens, ficheiros Docker, `docker-compose.yml`, `.env.example`, documentação ou repositório.

## Estrutura Simples E Profissional Do Repositório

A IA deve manter a estrutura do repositório simples, previsível e profissional.

A estrutura deve favorecer:

- leitura rápida;
- separação clara de responsabilidades;
- onboarding fácil;
- manutenção futura;
- compatibilidade com Docker, CI/CD e deploy;
- facilidade para outras IAs entenderem o projeto.

Evitar:

- pastas sem propósito claro;
- nomes vagos como `coisas`, `misc`, `old`, `teste_final`, `novo2`;
- duplicação desnecessária;
- ficheiros soltos sem organização;
- estruturas demasiado complexas para projetos simples;
- refactors grandes sem necessidade.

## Estrutura Recomendada Do Repositório

A estrutura deve ser adaptada ao projeto, mas seguir um padrão simples quando aplicável.

Para projetos simples:

```text
projeto/
├── AGENTS.md
├── PROJECT_CONTEXT.md
├── README.md
├── CHANGELOG.md
├── CHANGELOG_POLICY.md
├── .env.example
├── .gitignore
├── .dockerignore
├── docker-compose.yml
│
├── tasks/
│   ├── todo.md
│   └── lessons.md
│
├── docs/
│   └── ...
│
├── src/
│   └── ...
│
├── tests/
│   └── ...
│
└── scripts/
    └── ...
```

Para projetos full-stack:

```text
projeto/
├── backend/
├── frontend/
├── database/
├── docs/
├── scripts/
├── tests/
├── tasks/
├── docker-compose.yml
├── .env.example
├── README.md
├── PROJECT_CONTEXT.md
├── CHANGELOG.md
├── CHANGELOG_POLICY.md
└── AGENTS.md
```

## Critério De Simplicidade Estrutural

Antes de criar novas pastas, serviços, containers ou abstrações, a IA deve perguntar:

- isto reduz complexidade ou aumenta?
- isto ajuda no deploy, testes ou manutenção?
- isto é adequado à escala atual do projeto?
- uma pessoa nova conseguiria perceber a estrutura em menos de 5 minutos?

Se a resposta for negativa, manter a solução mais simples.

## Critério De Conclusão Para Documentação

Uma tarefa que altere comportamento, instalação, configuração, comandos, arquitetura, endpoints, scripts, migrations ou fluxos críticos só pode ser considerada concluída se:

- `README.md` estiver criado ou atualizado;
- a estrutura do repositório estiver simples, coerente e profissional;
- Docker estiver documentado no `README.md` quando for usado;
- badges do `README.md` estiverem corretos quando usados;
- `PROJECT_CONTEXT.md` estiver coerente com o projeto;
- `CHANGELOG.md` tiver entrada versionada quando aplicável;
- os comentários e documentação técnica estiverem em português europeu correto;
- a documentação não contradisser o código;
- a resposta final indicar claramente se o `README.md` foi atualizado ou se não foi necessário.

## Changelog Obrigatório

Cada projeto deve ter um `CHANGELOG.md` versionado na raiz.

A política completa fica em `CHANGELOG_POLICY.md` e é obrigatória para qualquer IA que trabalhe no repositório.

Regras mínimas:

- Atualizar `CHANGELOG.md` automaticamente sempre que houver alteração versionável.
- Usar SemVer: `MAJOR.MINOR.PATCH`.
- Inserir entradas novas no topo do ficheiro.
- Nunca apagar histórico antigo.
- Registar motivo, impacto, ficheiros alterados, testes, validação, refs e diff resumido.
- Não entregar trabalho como concluído se houve alteração versionável sem nova entrada no changelog.
- Se não houve alteração versionável, declarar isso na resposta final.

## Orquestração do Fluxo de Trabalho

### 1. Modo de Planeamento por Defeito

- Entrar em modo de planeamento para QUALQUER tarefa não trivial, ou seja, tarefas com 3 ou mais passos, decisões de arquitetura ou alterações com impacto transversal.
- Se algo correr mal, PARAR e replanear imediatamente — não continuar a forçar uma abordagem errada.
- Usar o modo de planeamento para passos de verificação, não apenas para construir.
- Escrever especificações detalhadas à partida para reduzir a ambiguidade.

### 2. Estratégia de Subagentes

- Usar subagentes de forma liberal para manter limpa a janela de contexto principal.
- Delegar investigação, exploração e análise paralela em subagentes.
- Para problemas complexos, aplicar mais capacidade computacional através de subagentes.
- Uma linha de abordagem por subagente para uma execução focada.

### 3. Ciclo de Autoaperfeiçoamento

- Após QUALQUER correção do utilizador, atualizar `tasks/lessons.md` com o padrão identificado.
- Escrever regras para prevenir o mesmo erro no futuro.
- Iterar sobre estas lições até a taxa de erros descer.
- Rever as lições no início da sessão para o projeto relevante.

### 4. Verificação Antes De Concluir

- Nunca marcar uma tarefa como concluída sem provar que funciona.
- Comparar o comportamento entre a `main` e as alterações atuais quando for relevante.
- Perguntar: “Um Staff Engineer aprovaria isto?”
- Executar testes, verificar logs e demonstrar correção.

### 5. Exigir Elegância De Forma Equilibrada

- Para alterações não triviais, pausar e perguntar: “há uma forma mais elegante?”
- Se uma correção parecer improvisada, implementar a solução elegante com base no conhecimento atual.
- Ignorar isto para correções simples e óbvias — não complicar em excesso.
- Questionar o próprio trabalho antes de o apresentar.

### 6. Correção Autónoma De Bugs

- Quando receber um relatório de bug, corrigir o problema. Não pedir acompanhamento passo a passo ao utilizador.
- Identificar logs, erros e testes a falhar — depois resolver.
- Não exigir troca de contexto ao utilizador quando o repositório tiver informação suficiente.
- Corrigir testes de CI a falhar sem que seja necessário explicar como.

## Gestão De Tarefas

1. **Planear Primeiro**: escrever o plano em `tasks/todo.md` com itens verificáveis.
2. **Verificar O Plano**: confirmar o plano antes de iniciar a implementação.
3. **Acompanhar O Progresso**: marcar itens como concluídos à medida que a execução avança.
4. **Explicar Alterações**: registar resumo de alto nível em cada passo relevante.
5. **Documentar Resultados**: adicionar uma secção de revisão a `tasks/todo.md`.
6. **Capturar Lições**: atualizar `tasks/lessons.md` após correções, erros ou feedback do utilizador.

## Princípios Fundamentais

- **Simplicidade Primeiro**: tornar cada alteração tão simples quanto possível. Impactar o mínimo de código.
- **Sem Preguiça**: encontrar causas raiz. Sem correções temporárias. Padrões de programador sénior.
- **Impacto Mínimo**: as alterações devem tocar apenas no que é necessário. Evitar introduzir bugs.
- **Documentação Clara**: todo o código relevante deve estar documentado em português europeu correto, com acentuação e sem ambiguidade.
- **README Profissional**: todo o projeto deve ter um `README.md` claro, útil, atualizado e visualmente apelativo, ao nível dos melhores repositórios.
- **Docker Adequado**: Docker deve ser usado quando trouxer reprodutibilidade, facilidade de deploy ou isolamento útil, sem criar complexidade desnecessária.
- **Estrutura Profissional**: o repositório deve manter uma organização simples, previsível, limpa e fácil de compreender.
- **Badges Reais**: os badges do `README.md` devem refletir apenas informação confirmada sobre stack, licença, versão, estado, testes ou CI/CD.
