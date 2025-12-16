# Sistema de Gestão de Férias para Google Sheets

Script desenvolvido por **Emanuel Ferreira** (@emanuwells) para automatizar a gestão de férias no Google Sheets, incluindo contadores automáticos e sincronização bidirecional com o Google Calendar.

## 🚀 Funcionalidades

- **Contagem Automática**: Calcula dias de férias gozados e planeados baseado na cor das células.
- **Sincronização com Calendar**: Cria eventos no Google Calendar para os dias marcados, evitando duplicados.
- **Deteção de Aniversário**: Identifica e gere o dia de aniversário da empresa (célula verde).
- **Agrupamento Inteligente**: Dias consecutivos são agrupados num único evento no calendário (ex: "Férias (5 dias)").
- **Automação**:
    - Atualização instantânea ao editar valores.
    - Sincronização automática a cada 5 minutos (opcional).

## 🛠️ Configuração

1. **Abra o seu Google Sheet de Férias**.
2. **Extensions > Apps Script**: Cole o código do ficheiro `Vacation_Mode.js`.
3. **Ajuste as Configurações** (no início do ficheiro):
    ```javascript
    const CONFIG = {
      CALENDAR_RANGE: 'G5:AI16', // Área onde pinta os dias
      CORES: {
        FERIAS_ATUAL: '#d9d2e9', // Cor das férias
        ANIVERSARIO: '#d9ead3'   // Cor do aniversário
      },
      // ... outras configurações
    };
    ```
4. **Execute `configurarSheet`**: Selecione esta função e execute-a uma vez para criar a legenda e os contadores automaticamente.

## 🎨 Cores Padrão

- **Roxo (#d9d2e9)**: Férias planeadas/gozadas.
- **Verde (#d9ead3)**: Dia de aniversário.
- **Amarelo (#fff2cc)**: Férias transitadas do ano anterior.

## 📋 Menu Personalizado

O script cria um menu **"🏖️ Gestão de Férias"** no seu Sheet com as opções:
- **Sincronizar Tudo**: Atualiza contadores e envia para o Calendar.
- **Ativar Automação**: Liga os triggers automáticos.
- **Diagnóstico**: Ferramentas para verificar se as cores estão a ser detetadas corretamente.


## 📄 Licença

Este projeto é de uso livre. Atribuição ao autor original é apreciada.

## 🔗 Créditos e Referências

O template Excel base utilizado neste projeto foi adaptado a partir do **Calendário Excel com Feriados – Portugal** disponível em [economiafinancas.com](https://economiafinancas.com/2025/), com personalizações e alterações para integração com este script.
