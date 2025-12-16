# Manual de Utilização - Sistema de Gestão de Férias

**Autor:** Emanuel Ferreira
**Versão:** 1.3.0
**Data:** Dezembro 2025

---

## 1. Introdução

Este sistema foi desenvolvido para facilitar a gestão de férias em equipas ou uso pessoal, utilizando o Google Sheets e o Google Calendar. O objetivo é permitir que o utilizador marque visualmente os dias de férias numa folha de cálculo e o sistema trate automaticamente da contagem de dias e da criação de eventos no calendário.

### Principais Benefícios
*   **Visual:** Marcação de férias por cores.
*   **Automático:** Contagem de dias gozados vs. planeados.
*   **Integrado:** Sincronização direta com o Google Calendar.
*   **Inteligente:** Agrupa dias seguidos num único evento (ex: 15 a 30 de Agosto).

---

## 2. Instalação

Como este é um script para Google Sheets, a instalação é feita dentro do próprio documento.

1.  Abra o seu ficheiro Google Sheets.
2.  No menu superior, vá a **Extensões** > **Apps Script**.
3.  Apague qualquer código que lá esteja.
4.  Copie o código do ficheiro `Vacation_Mode.js` e cole no editor.
5.  Clique no ícone de **Guardar** (disquete).
6.  Atualize a página do seu Google Sheet (F5).
7.  Deverá aparecer um novo menu chamado **"🏖️ Gestão de Férias"**.

---

## 3. Configuração Inicial

Antes de começar a usar, é necessário configurar o sistema para o seu layout específico.

### 3.1 Ajustar o Script
No topo do código (no Apps Script), encontrará uma secção chamada `CONFIG`. Ajuste os valores conforme necessário:

*   **`CALENDAR_RANGE`**: Defina o intervalo de células onde está o seu calendário (Ex: `'G5:AI16'`).
*   **`CORES`**: Confirme se os códigos hexadecimais correspondem às cores que vai usar. O padrão é:
    *   Roxo para Férias (`#d9d2e9`)
    *   Verde para Aniversário (`#d9ead3`)

### 3.2 Preparar a Folha
1.  Vá ao menu **"🏖️ Gestão de Férias"**.
2.  (Na primeira utilização, o Google pedirá perimssão para executar o script. Aceite todas as permissões).
3.  O script funciona melhor se tiver uma área de legenda/contadores. Pode usar a estrutura que o script cria automaticamente se desejar, mas certifique-se que o código aponta para as células certas em `CONFIG.CELULAS`.

---

## 4. Como Usar

### 4.1 Marcar Férias
Basta selecionar as células correspondentes aos dias desejados e alterar a **Cor de Fundo**:
*   Use **Roxo** para dias de férias.
*   Use **Verde** para o dia de aniversário.

**Nota:** O sistema ignora células que não tenham números (dias), por isso pode pintar linhas inteiras sem problema.

### 4.2 Sincronizar
Existem duas formas de atualizar os dados:

**Método Manual:**
1.  Terminou de marcar as férias?
2.  Clique no menu **"🏖️ Gestão de Férias"** > **"⚡ SINCRONIZAR TUDO"**.
3.  Aguarde a mensagem de sucesso. Os contadores serão atualizados e os eventos aparecerão no seu Google Calendar.

**Método Automático:**
1.  No menu, selecione **"🤖 Ativar Sincronização Automática"**.
2.  A partir de agora, o sistema verifica a cada 5 minutos se houve alterações e atualiza tudo sozinho.

---

## 5. Resolução de Problemas

### 5.1 O calendário não atualiza
*   Verifique se as cores que usou são exatamente as mesmas definidas no código.
*   Use a opção de menu **"🔍 Testar Deteção de Cores"** para ver o que o sistema está a "ver".

### 5.2 Erro de Permissões
*   Se o script falhar ao aceder ao calendário, verifique se a conta Google que está a usar é a mesma onde quer criar os eventos.
*   Tente executar o script novamente e valide se aceitou todas as permissões.

### 5.3 Eventos Duplicados
*   O sistema remove automaticamente eventos antigos criados por ele antes de criar novos. Se vir duplicados, pode ser porque o título ou a assinatura do evento foi alterada manualmente no Calendar.
*   Recomenda-se não editar os eventos criados pelo script manualmente no Calendar; faça as alterações no Sheet e sincronize novamente.

---


## 6. Créditos

Este script foi desenvolvido para ser genérico e flexível.
**Autor Original:** Emanuel Ferreira
**Contacto:** @emanuwells
**Base do Calendário:** Adaptado de [economiafinancas.com](https://economiafinancas.com/2025/) (Calendário Excel com Feriados – Portugal).
