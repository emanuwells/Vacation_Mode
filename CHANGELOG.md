# 🏖️ Vacation Mode — Google Sheets + Calendar

Este script transforma o teu planeamento de férias no **Google Sheets** (feito através de células pintadas) em contagens automáticas e eventos sincronizados no **Google Calendar**. Suporta múltiplos anos e automatização total via triggers.

---

## ✨ Funcionalidades Principais

* 📅 **Gestão de Datas:** Conta automaticamente dias de férias (gozados/planeados) e aniversários.
* 🗓️ **Sincronização Inteligente:** Cria eventos no Google Calendar sem duplicados, agrupando dias consecutivos (ex: 5 dias de férias = 1 evento longo).
* 🔄 **Suporte Multi-Ano:** Deteta e processa todas as folhas que sigam o padrão `Calendario YYYY` ou `Calendário YYYY`.
* 🧭 **Interface Nativa:** Adiciona um menu personalizado ao Google Sheets para ações rápidas.
* ⏱️ **Automação:** Opção para ativar uma sincronização automática a cada 5 minutos.

---

## 🚀 Instalação Rápida

1. No teu Google Sheets, vai a **Extensões** → **Apps Script**.
2. Apaga qualquer código existente no editor.
3. Copia e cola o conteúdo do ficheiro `Vacation_Mode.js`.
4. Guarda o projeto (`Ctrl+S`) e faz **Refresh (F5)** na folha de cálculo.
5. O menu **"🌴 Gestão de Férias"** aparecerá na barra superior.

---

## ⚙️ Configuração do Script

No topo do ficheiro `Vacation_Mode.js`, encontrarás o objeto `CONFIG`. Ajusta os valores conforme a tua estrutura:

| Variável | Descrição | Valor Padrão |
| :--- | :--- | :--- |
| `CALENDAR_RANGE` | Intervalo da grelha do calendário. | `'G5:AI16'` |
| `CORES` | Códigos hexadecimais para Férias e Aniversário. | *(Configurável)* |
| `CELULAS` | Onde o script deve escrever os resultados/contadores. | *(Ajustar à legenda)* |
| `CALENDARIO.NOME` | Nome do calendário (vazio = Calendário Principal). | `''` |
| `TITULO_EVENTO` | Prefixo do nome do evento no calendário. | `'Férias'` |

### 📂 Estrutura das Folhas
Para que o script funcione, nomeia as tuas abas como:
* `Calendario 2025`
* `Calendário 2026`
* *(Dica: Podes simplesmente duplicar a folha de um ano para o outro.)*

---

## 🕹️ Como Utilizar

### Modo Manual (Recomendado)
1. Pinta os dias de férias ou aniversário na grelha do Sheets com as cores definidas.
2. Vai ao menu **🌴 Gestão de Férias** → **🔄 SINCRONIZAR TUDO**.
3. O script atualizará os contadores na folha e criará/removerá os eventos no Calendar.

### Modo Automático
1. No menu, seleciona **Ativar Sincronização Automática**.
2. O script criará um *trigger* que corre de 5 em 5 minutos para manter tudo atualizado sem intervenção manual.

---

## 💡 Dicas e Resolução de Problemas

> **Dica sobre Cores:**
> Cada monitor ou tema pode alterar ligeiramente a perceção da cor. Se o script não detetar as tuas marcações, usa a opção **"Diagnóstico: Testar Deteção de Cores"** no menu para confirmar o código hexadecimal exato que o Google Sheets está a ler.

* **Eventos não aparecem:** Verifica se o `CALENDAR_RANGE` cobre todos os dias do mês na tua folha.
* **Permissões:** Na primeira execução, o Google pedirá autorização para aceder ao Sheets e ao Calendar. É um processo seguro e necessário.
* **Calendário Alvo:** Se usares um calendário partilhado, garante que tens permissões de edição e que o nome em `CALENDARIO.NOME` é exatamente igual ao que vês no Google Calendar.

---

## 🛠️ Desenvolvimento e Licença

* **Versão Atual:** 1.3.2
* **Autor:** Emanuel Ferreira (@emanuwells)
* **Licença:** MIT (Atribuição apreciada)