# Vacation Mode – Google Sheets + Calendar

Script desenvolvido para automatizar a gestão de férias no Google Sheets, incluindo contadores automáticos e sincronização direta com o Google Calendar. Suporta vários anos em paralelo.

Baseado no Calendário em Excel com Feriados, disponível em https://economiafinancas.com/.

## O que faz
- 📆 Conta dias de férias gozados e planeados, bem como o dia de aniversário (cores configuráveis).
- 🗓️ Cria eventos no Google Calendar sem duplicados, agrupando dias consecutivos num único evento.
- 🔁 Multi-ano: percorre todas as folhas cujo nome contenha `Calendario YYYY` ou `Calendário YYYY`.
- 🧭 Menu no Sheets com ações rápidas: sincronização completa, triggers e diagnóstico de cores.
- ⏱️ Trigger opcional de 5 minutos para sincronização automática.

## Instalação
1. No Google Sheets, aceda a Extensões → Apps Script.
2. Apague o código existente e cole o conteúdo de `Vacation_Mode.js`.
3. Guarde e volte ao Sheet (F5). O menu “Gestão de Férias” aparece.

### Configurar `CONFIG` (topo do ficheiro)
- `CALENDAR_RANGE`: intervalo do calendário; por omissão, `C5:AM16`.
- `CORES`: cores usadas para férias e aniversário.
- `CELULAS`: células onde estão os contadores; ajuste se a legenda estiver noutro sítio.
- `CALENDARIO.NOME`: deixe vazio para usar o calendário principal, ou defina o nome exato de um calendário que possua.
- `CALENDARIO.TITULO_EVENTO`: título-base dos eventos, por exemplo, "Férias".

### Estrutura das folhas
- Crie ou renomeie folhas como `Calendario 2025`, `Calendario 2026`, etc. (com ou sem acento).
- Use a mesma grelha em cada folha; basta duplicar a folha para o ano seguinte e pintar.

## Como usar
### Manual (recomendado)
1. Pinte os dias de férias ou de aniversário na(s) folha(s).
2. Aceda a “Gestão de Férias” → “SINCRONIZAR TUDO”.
3. Contadores e eventos de todas as folhas de calendário são atualizados.

### Automático (opcional)
1. Aceda a “Gestão de Férias” → “Ativar Sincronização Automática”.
2. O script corre quando pinta células e, como salvaguarda, a cada 5 minutos: conta, sincroniza e escreve no Calendar.

## Dicas e resolução de problemas
- Sem eventos criados: confirme que as cores usadas correspondem a `CONFIG.CORES`. Use “Testar Deteção de Cores” no menu.
- Eventos a desaparecer: o script só limpa eventos depois de encontrar células pintadas; garanta nomes `Calendario YYYY` e intervalo/cores corretos.
- Calendário-alvo: deixe `CALENDARIO.NOME` vazio para usar o principal, ou indique um calendário que seja seu.

## Desenvolvimento
- Ficheiro principal: `Vacation_Mode.js`.
- Versão: 1.3.3.
- Changelog: `CHANGELOG.md`.

## Licença
MIT. Atribuição apreciada: Emanuel Ferreira (@emanuwells).
