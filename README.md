# Vacation Mode ? Google Sheets + Calendar

Script em Google Apps Script para transformar um calend?rio pintado no Google Sheets em contagem autom?tica de f?rias/anivers?rio e eventos no Google Calendar. Suporta v?rios anos ao mesmo tempo.

## O que faz
- Conta dias de f?rias gozados/planeados e o dia de anivers?rio (cores configur?veis).
- Cria eventos no Google Calendar sem duplicados, agrupando dias consecutivos num ?nico evento.
- Multi-ano: percorre todas as folhas cujo nome contenha `Calendario YYYY` ou `Calend?rio YYYY`.
- Menu no Sheets com a??es r?pidas (Sincronizar Tudo, triggers, diagn?stico de cores).
- Trigger opcional de 5 minutos para sincroniza??o autom?tica.

## Instala??o
1. No Google Sheets: Extens?es ? Apps Script.
2. Apague o c?digo existente e cole o conte?do de `Vacation_Mode.js`.
3. Guarde e volte ao Sheet (F5). O menu ?Gest?o de F?rias? aparece.

### Configurar `CONFIG` (topo do ficheiro)
- `CALENDAR_RANGE`: intervalo do calend?rio (padr?o `G5:AI16`).
- `CORES`: cores usadas para f?rias e anivers?rio.
- `CELULAS`: c?lulas onde est?o os contadores (ajuste se a legenda estiver noutro s?tio).
- `CALENDARIO.NOME`: deixe vazio para usar o calend?rio principal ou defina o nome exato de um calend?rio que possua.
- `CALENDARIO.TITULO_EVENTO`: t?tulo base dos eventos (ex.: "F?rias").

### Estrutura das folhas
- Crie/renomeie folhas como `Calendario 2025`, `Calendario 2026`, etc. (com ou sem acento).
- Use a mesma grelha em cada folha; basta duplicar a folha para o ano seguinte e pintar.

## Como usar
### Manual (recomendado)
1. Pinte os dias de f?rias/anivers?rio na(s) folha(s).
2. Menu ?Gest?o de F?rias? ? ?SINCRONIZAR TUDO?.
3. Contadores e eventos de todas as folhas de calend?rio s?o atualizados.

### Autom?tico (opcional)
1. Menu ?Gest?o de F?rias? ? ?Ativar Sincroniza??o Autom?tica (5 min)?.
2. O script corre a cada 5 minutos: conta, sincroniza e escreve no Calendar.

## Dicas e resolu??o de problemas
- Sem eventos criados: confirme que as cores usadas batem com `CONFIG.CORES`. Use ?Testar Dete??o de Cores? no menu.
- Eventos a desaparecer: o script s? limpa eventos depois de encontrar c?lulas pintadas; garanta nomes `Calendario YYYY` e range/cores corretos.
- Calend?rio alvo: deixe `CALENDARIO.NOME` vazio para usar o principal ou indique um calend?rio que seja seu (owned).

## Desenvolvimento
- Ficheiro principal: `Vacation_Mode.js`.
- Vers?o: 1.3.2.
- Changelog: `CHANGELOG.md`.

## Licen?a
MIT. Atribui??o apreciada: Emanuel Ferreira (@emanuwells).
