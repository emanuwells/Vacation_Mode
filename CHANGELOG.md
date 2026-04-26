# Changelog

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
