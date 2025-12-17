# Vacation Mode – Google Sheets + Calendar

Script em Google Apps Script para transformar um calendário pintado no Google Sheets em contagem automática de férias/aniversário e eventos no Google Calendar. Suporta vários anos ao mesmo tempo.

## O que faz
- 📆 Conta dias de férias gozados/planeados e o dia de aniversário (cores configuráveis).
- 🗓️ Cria eventos no Google Calendar sem duplicados, agrupando dias consecutivos num único evento.
- 🔁 Multi-ano: percorre todas as folhas cujo nome contenha `Calendario YYYY` ou `Calendário YYYY`.
- 🧭 Menu no Sheets com ações rápidas (Sincronizar Tudo, triggers, diagnóstico de cores).
- ⏱️ Trigger opcional de 5 minutos para sincronização automática.

## Instalação
1. No Google Sheets: Extensões → Apps Script.
2. Apague o código existente e cole o conteúdo de `Vacation_Mode.js`.
3. Guarde e volte ao Sheet (F5). O menu “Gestão de Férias” aparece.

### Configurar `CONFIG` (topo do ficheiro)
- `CALENDAR_RANGE`: intervalo do calendário (padrão `G5:AI16`).
- `CORES`: cores usadas para férias e aniversário.
- `CELULAS`: células onde estão os contadores (ajuste se a legenda estiver noutro sítio).
- `CALENDARIO.NOME`: deixe vazio para usar o calendário principal ou defina o nome exato de um calendário que possua.
- `CALENDARIO.TITULO_EVENTO`: título base dos eventos (ex.: "Férias").

### Estrutura das folhas
- Crie/renomeie folhas como `Calendario 2025`, `Calendario 2026`, etc. (com ou sem acento).
- Use a mesma grelha em cada folha; basta duplicar a folha para o ano seguinte e pintar.

## Como usar
### Manual (recomendado)
1. Pinte os dias de férias/aniversário na(s) folha(s).
2. Menu “Gestão de Férias” → “SINCRONIZAR TUDO”.
3. Contadores e eventos de todas as folhas de calendário são atualizados.

### Automático (opcional)
1. Menu “Gestão de Férias” → “Ativar Sincronização Automática (5 min)”.
2. O script corre a cada 5 minutos: conta, sincroniza e escreve no Calendar.

## Dicas e resolução de problemas
- Sem eventos criados: confirme que as cores usadas batem com `CONFIG.CORES`. Use “Testar Deteção de Cores” no menu.
- Eventos a desaparecer: o script só limpa eventos depois de encontrar células pintadas; garanta nomes `Calendario YYYY` e range/cores corretos.
- Calendário alvo: deixe `CALENDARIO.NOME` vazio para usar o principal ou indique um calendário que seja seu (owned).

## Desenvolvimento
- Ficheiro principal: `Vacation_Mode.js`.
- Versão: 1.3.2.
- Changelog: `CHANGELOG.md`.

## Licença
MIT. Atribuição apreciada: Emanuel Ferreira (@emanuwells).
