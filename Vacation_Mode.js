/**
 * SISTEMA DE GESTAO DE FERIAS
 * Versao: 1.3.2
 * Data: 2025-12-17
 * 
 * Autor: Emanuel Ferreira (@emanuwells)
 * 
 * Descrição:
 * Sistema completo para gestão de férias com:
 * - Contagem automática de dias de férias (células roxas)
 * - Contagem automática de dia de aniversário (células verdes)
 * - Sincronização com Google Calendar (com agrupamento de dias consecutivos)
 * - Atualização automática via trigger
 * - Menu personalizado
 * 
 */

// ============================
// CONFIGURAÇÕES GLOBAIS
// ============================

const CONFIG = {
  // Range do calendário (12 linhas = meses, 31 colunas = dias)
  // Ajusta aqui se mudares a posição do quadro; no teu layout o topo do calendário começa em G5 e o dia 31 cai em AI16.
  CALENDAR_RANGE: 'G5:AI16',

  // Cores a detetar (hexadecimal - Google Sheets format)
  CORES: {
    FERIAS_ATUAL: '#d9d2e9',      // Roxo/lilás - férias planeadas/ano corrente
    FERIAS_ATUAL_ALT: '#b4a7d6',  // Variante de roxo/lilás comum em folhas
    FERIAS_ANTERIOR: '#fff2cc',   // Amarelo claro - dias transitados do ano anterior
    ANIVERSARIO: '#d9ead3'        // Verde claro - dia de aniversário
  },

  // Células onde aparecem os contadores (coluna B)
  CELULAS: {
    // Contadores de férias
    FERIAS_DISPONIVEIS: 'C18',      // Input manual do utilizador (ano corrente)
    FERIAS_ANTERIOR: 'C19',         // Dias transitados do ano anterior (input manual)
    FERIAS_GOZADAS: 'C20',          // Calculado automaticamente
    FERIAS_PLANEADAS: 'C21',        // Calculado automaticamente
    FERIAS_TOTAL: 'C22',            // Soma: gozadas + planeadas
    FERIAS_RESTANTES: 'C23',        // Diferença: disponíveis + anteriores - total

    // Contadores de aniversário
    ANIVERSARIO_DISPONIVEL: 'C25', // Fixo: 1 dia
    ANIVERSARIO_GOZADO: 'C26',     // 0 ou 1 (se data passou)
    ANIVERSARIO_A_GOZAR: 'C27'     // 0 ou 1 (se data futura)
  },

  // Configurações do Google Calendar
  CALENDARIO: {
    NOME: '',                      // Deixe vazio para usar o Calendário Principal (recomendado)
    TITULO_EVENTO: 'Férias',       // Título dos eventos criados
    ANO: new Date().getFullYear(), // Ano padrão (é substituído pelo ano da folha se existir)
    MARCADOR: '[FERIAS_AUTO]'      // Assinatura para identificar eventos gerados pelo script
  }
};

/**
 * Deteta o ano da folha pelo nome (ex.: "Calendário 2025", "Calendario 2026").
 * Se não encontrar, devolve o ano padrão configurado.
 */
function obterAnoDaSheet(sheet) {
  const nome = sheet.getName();
  const match = nome.match(/20\d{2}/);
  return match ? parseInt(match[0], 10) : CONFIG.CALENDARIO.ANO;
}

/**
 * Devolve as folhas de calendário (nome começando por "Calendario"/"Calendário" e ano YYYY).
 * Se nenhuma for encontrada, devolve apenas a folha ativa como fallback.
 */
function obterFolhasCalendario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const todas = ss.getSheets();

  const alvo = todas
    .map(sheet => ({ sheet, ano: obterAnoDaSheet(sheet) }))
    .filter(({ sheet }) => /Calend.?rio\s*\d{4}/i.test(sheet.getName()));

  if (alvo.length === 0) {
    const ativa = ss.getActiveSheet();
    return [{ sheet: ativa, ano: obterAnoDaSheet(ativa) }];
  }

  return alvo;
}

// ============================
// ============================
// FUNÇÃO PRINCIPAL - ATUALIZAR CONTADORES
// ============================

/**
 * Atualiza todos os contadores de férias e aniversário
 * Conta células coloridas e distingue entre datas passadas e futuras
 */
function atualizarContadores(e, sheetParam, anoParam) {
  try {
    const sheet = sheetParam || (e && e.range ? e.range.getSheet() : SpreadsheetApp.getActiveSpreadsheet().getActiveSheet());
    const ano = anoParam || obterAnoDaSheet(sheet);
    const hoje = obterDataHoje();

    // Obter dados do calend?rio (valores e cores de fundo)
    const range = sheet.getRange(CONFIG.CALENDAR_RANGE);
    const valores = range.getValues();
    const cores = range.getBackgrounds();

    // Inicializar contadores
    const contadores = {
      feriasGozadas: 0,
      feriasPlaneadas: 0,
      aniversarioGozado: 0,
      aniversarioAGozar: 0
    };

    // Percorrer todas as c?lulas do calend?rio
    for (let linha = 0; linha < valores.length; linha++) {
      for (let coluna = 0; coluna < valores[linha].length; coluna++) {
        processarCelula(valores[linha][coluna], cores[linha][coluna], linha, hoje, contadores, ano);
      }
    }

    // Atualizar c?lulas no sheet com os novos valores
    atualizarCelulasContadores(sheet, contadores);

    Logger.log('? Contadores atualizados com sucesso! (' + sheet.getName() + ' - ' + ano + ')');
    mostrarNotificacao(sheet.getName() + ': Contadores atualizados!', 'Sucesso', 3);

  } catch (erro) {
    Logger.log('? Erro ao atualizar contadores: ' + erro.message);
    mostrarNotificacao('Erro ao atualizar contadores. Verifica o log.', 'Erro', 5);
  }
}

/**
 * Obtém a data de hoje normalizada (sem horas)
 */
function obterDataHoje() {
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  return hoje;
}

/**
 * Processa uma célula individual e atualiza os contadores
 */
function processarCelula(valor, cor, indiceLinha, hoje, contadores, ano) {
  // Ignorar c?lulas vazias ou com texto (S, D, T, Q)
  if (!valor || isNaN(valor)) {
    return;
  }

  // Construir a data completa (ano derivado da folha), m?s baseado no ?ndice da linha, dia = valor da c?lula
  const mes = indiceLinha; // 0=Janeiro, 1=Fevereiro, ..., 11=Dezembro
  const dia = parseInt(valor);
  const data = new Date(ano, mes, dia);

  // Normalizar cor (min?sculas, sem espa?os)
  const corNormalizada = cor.toLowerCase().replace(/\s/g, '');

  // Verificar se ? c?lula de f?rias (aceita cor do ano corrente ou do ano anterior)
  const corFeriasAtual = CONFIG.CORES.FERIAS_ATUAL.toLowerCase();
  const corFeriasAtualAlt = CONFIG.CORES.FERIAS_ATUAL_ALT.toLowerCase();
  const corFeriasAnterior = CONFIG.CORES.FERIAS_ANTERIOR.toLowerCase();
  const isFerias =
    corNormalizada === corFeriasAtual ||
    corNormalizada === corFeriasAtualAlt ||
    corNormalizada === corFeriasAnterior;

  if (isFerias) {
    if (data <= hoje) {
      contadores.feriasGozadas++;
    } else {
      contadores.feriasPlaneadas++;
    }
  }

  // Verificar se ? c?lula de anivers?rio (verde)
  if (corNormalizada === CONFIG.CORES.ANIVERSARIO.toLowerCase()) {
    if (data <= hoje) {
      contadores.aniversarioGozado = 1; // M?ximo 1 dia
    } else {
      contadores.aniversarioAGozar = 1; // M?ximo 1 dia
    }
  }
}

function atualizarCelulasContadores(sheet, contadores) {
  // Atualizar contadores de férias
  sheet.getRange(CONFIG.CELULAS.FERIAS_GOZADAS).setValue(contadores.feriasGozadas);
  sheet.getRange(CONFIG.CELULAS.FERIAS_PLANEADAS).setValue(contadores.feriasPlaneadas);

  const totalPlaneado = contadores.feriasGozadas + contadores.feriasPlaneadas;
  sheet.getRange(CONFIG.CELULAS.FERIAS_TOTAL).setValue(totalPlaneado);

  // Calcular e atualizar dias restantes
  const disponiveis = sheet.getRange(CONFIG.CELULAS.FERIAS_DISPONIVEIS).getValue() || 0;
  const anterior = sheet.getRange(CONFIG.CELULAS.FERIAS_ANTERIOR).getValue() || 0;
  const restantes = (disponiveis + anterior) - totalPlaneado;
  sheet.getRange(CONFIG.CELULAS.FERIAS_RESTANTES).setValue(restantes);

  // Atualizar contadores de aniversário
  sheet.getRange(CONFIG.CELULAS.ANIVERSARIO_GOZADO).setValue(contadores.aniversarioGozado);
  sheet.getRange(CONFIG.CELULAS.ANIVERSARIO_A_GOZAR).setValue(contadores.aniversarioAGozar);
}

// ============================
// SINCRONIZAÇÃO COMPLETA (CONTADORES + CALENDAR)
// ============================

/**
 * Sincroniza tudo: atualiza contadores E sincroniza com Google Calendar
 * Função combinada para usar com botão ou trigger automático
 */
function sincronizarTudo() {
  try {
    Logger.log('?? Iniciando sincroniza??o completa...');

    const folhas = obterFolhasCalendario();
    if (folhas.length === 0) {
      Logger.log('?? Nenhuma folha de calend?rio encontrada.');
      mostrarNotificacao('Nenhuma folha de calend?rio encontrada.', 'Aviso', 5);
      return;
    }

    folhas.forEach(({ sheet, ano }) => {
      atualizarContadores(null, sheet, ano);
      Utilities.sleep(500);
      sincronizarComCalendar(sheet, ano);
    });

    Logger.log('? Sincroniza??o completa finalizada!');
    mostrarNotificacao('Contadores e Calendar sincronizados!', 'Sincroniza??o Completa', 5);

  } catch (erro) {
    Logger.log('? Erro na sincroniza??o completa: ' + erro.message);
    mostrarNotificacao('Erro na sincroniza??o. Verifica o log.', 'Erro', 5);
  }
}

/**
 * Sincroniza uma folha específica com o Calendar (usa o ano detetado na folha).
 */
function sincronizarComCalendar(sheetParam, anoParam) {
  let lock;
  try {
    const sheet = sheetParam || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const ano = anoParam || obterAnoDaSheet(sheet);

    lock = LockService.getDocumentLock();
    if (!lock.tryLock(30000)) {
      Logger.log('Outra sincroniza??o est? a correr. Abortado para evitar duplicados.');
      mostrarNotificacao('Outra sincroniza??o em curso. Tenta novamente em instantes.', 'Aviso', 4);
      return;
    }
    Logger.log('?? Iniciando sincroniza??o com Google Calendar para ' + sheet.getName() + ' (' + ano + ')...');

    // Obter ou aceder ao calend?rio
    const calendario = obterCalendario();
    if (!calendario) {
      Logger.log('? Erro: Calend?rio n?o encontrado');
      mostrarNotificacao('Erro ao aceder ao calend?rio. Verifica as permiss?es.', 'Erro', 5);
      return;
    }

    Logger.log('? Calend?rio obtido: ' + calendario.getName());

    // Obter todas as datas de f?rias do calend?rio
    const datasFerias = obterDatasFerias(sheet, ano);
    const feriasRestantes = sheet.getRange(CONFIG.CELULAS.FERIAS_RESTANTES).getValue() || 0;
    if (datasFerias.length === 0) {
      Logger.log('?? Nenhuma c?lula roxa de f?rias encontrada em ' + sheet.getName());
      mostrarNotificacao(sheet.getName() + ': Nenhum dia de f?rias encontrado para sincronizar.', 'Aviso', 3);
      return;
    }

    Logger.log('Total de dias de ferias encontrados (' + sheet.getName() + '): ' + datasFerias.length);

    // Limpar eventos antigos (evitar duplicados) apenas ap?s confirmar que h? dados a recriar
    limparEventosAntigos(calendario, ano);

    // Agrupar datas consecutivas em blocos
    const blocos = agruparDatasConsecutivas(datasFerias);

    Logger.log('Agrupados em ' + blocos.length + ' periodo(s) de ferias');

    // Criar eventos no Calendar para cada bloco
    let eventosAdicionados = 0;
    blocos.forEach((bloco, index) => {
      try {
        const dataInicio = bloco.inicio;
        const dataFim = new Date(bloco.fim);
        dataFim.setDate(dataFim.getDate() + 1); // Calendar API precisa do dia seguinte para all-day events

        const numDias = bloco.dias.length;
        const titulo = numDias === 1
          ? CONFIG.CALENDARIO.TITULO_EVENTO
          : CONFIG.CALENDARIO.TITULO_EVENTO + ' (' + numDias + ' dias)';

        // Descrição com emojis, resumo e link do sheet para referência
        const resumoPeriodo = '📅 Período: ' + formatarData(dataInicio) + ' a ' + formatarData(bloco.fim) +
          ' (' + numDias + ' dias)';
        const resumoRestantes = '📉 Restantes: ' + feriasRestantes + ' dias';
        const linkSheet = '🔗 Sheet: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl();
        const descricaoEvento = [
          resumoPeriodo,
          resumoRestantes,
          linkSheet,
          CONFIG.CALENDARIO.MARCADOR
        ].join('\n');
        // Remover qualquer evento existente no mesmo periodo antes de criar (refor?o contra duplicados)
        const duplicados = calendario.getEvents(dataInicio, dataFim);
        duplicados.forEach(evento => {
          const tituloExistente = evento.getTitle();
          const descExistente = evento.getDescription() || '';
          const geradoPeloScript =
            tituloExistente.startsWith(CONFIG.CALENDARIO.TITULO_EVENTO) ||
            descExistente.indexOf(CONFIG.CALENDARIO.MARCADOR) !== -1;
          if (geradoPeloScript) {
            evento.deleteEvent();
            Logger.log('Removido duplicado antes de criar novo: ' + tituloExistente);
          }
        });

        calendario.createAllDayEvent(titulo, dataInicio, dataFim, { description: descricaoEvento });
        eventosAdicionados++;

        Logger.log('Bloco ' + (index + 1) + ': ' + formatarData(dataInicio) + ' a ' + formatarData(bloco.fim) + ' (' + numDias + ' dia(s))');

      } catch (erroEvento) {
        Logger.log('Erro ao criar evento: ' + erroEvento.message);
      }
    });

    const mensagem = eventosAdicionados === 1
      ? '1 periodo de ferias adicionado ao Google Calendar!'
      : eventosAdicionados + ' periodos de ferias adicionados ao Google Calendar!';

    Logger.log(eventosAdicionados + ' evento(s) criado(s) com sucesso');
    mostrarNotificacao(mensagem, 'Sincroniza??o completa', 5);

  } catch (erro) {
    Logger.log('? Erro ao sincronizar com Calendar: ' + erro.message);
    Logger.log('Stack trace: ' + erro.stack);
    mostrarNotificacao('Erro ao sincronizar. Verifica o log.', 'Erro', 5);
  } finally {
    if (lock) {
      lock.releaseLock();
    }
  }
}

function obterDatasFerias(sheet, ano) {
  const range = sheet.getRange(CONFIG.CALENDAR_RANGE);
  const valores = range.getValues();
  const cores = range.getBackgrounds();
  const startRow = range.getRow();      // linha real da 1? c?lula do calend?rio
  const startCol = range.getColumn();   // coluna real da 1? c?lula do calend?rio

  const datas = [];

  // Percorrer todas as c?lulas
  for (let linha = 0; linha < valores.length; linha++) {
    for (let coluna = 0; coluna < valores[linha].length; coluna++) {
      const valor = valores[linha][coluna];
      const cor = cores[linha][coluna];

      // Ignorar c?lulas vazias ou com texto
      if (!valor || isNaN(valor)) {
        continue;
      }

      // Normalizar cor
      const corNormalizada = cor.toLowerCase().replace(/\s/g, '');

      // Verificar se ? c?lula de f?rias (roxo)
      const corFeriasAtual = CONFIG.CORES.FERIAS_ATUAL.toLowerCase();
      const corFeriasAtualAlt = CONFIG.CORES.FERIAS_ATUAL_ALT.toLowerCase();
      const corFeriasAnterior = CONFIG.CORES.FERIAS_ANTERIOR.toLowerCase();
      const isFerias =
        corNormalizada === corFeriasAtual ||
        corNormalizada === corFeriasAtualAlt ||
        corNormalizada === corFeriasAnterior;

      if (isFerias) {
        const mes = linha; // 0=Janeiro, 1=Fevereiro, ..., 11=Dezembro
        const dia = parseInt(valor);
        const data = new Date(ano, mes, dia);

        datas.push(data);
        Logger.log('Encontrada: ' + formatarData(data) + ' (linha ' + (startRow + linha) + ', coluna ' + (startCol + coluna) + ')');
      }
    }
  }

  // Ordenar datas cronologicamente
  datas.sort((a, b) => a - b);

  return datas;
}

function agruparDatasConsecutivas(datas) {
  if (datas.length === 0) {
    return [];
  }

  const blocos = [];
  let blocoAtual = {
    inicio: datas[0],
    fim: datas[0],
    dias: [datas[0]]
  };

  for (let i = 1; i < datas.length; i++) {
    const dataAtual = datas[i];
    const dataAnterior = datas[i - 1];

    // Calcular diferença em dias
    const diferencaDias = Math.round((dataAtual - dataAnterior) / (1000 * 60 * 60 * 24));

    Logger.log('Comparando ' + formatarData(dataAnterior) + ' com ' + formatarData(dataAtual) + ': diferenca = ' + diferencaDias + ' dia(s)');

    if (diferencaDias === 1) {
      // Dias consecutivos - adicionar ao bloco atual
      blocoAtual.fim = dataAtual;
      blocoAtual.dias.push(dataAtual);
      Logger.log('Consecutivo! Bloco agora tem ' + blocoAtual.dias.length + ' dia(s)');
    } else {
      // Não consecutivo - fechar bloco atual e iniciar novo
      blocos.push(blocoAtual);
      Logger.log('Nao consecutivo! Fechando bloco de ' + blocoAtual.dias.length + ' dia(s)');

      blocoAtual = {
        inicio: dataAtual,
        fim: dataAtual,
        dias: [dataAtual]
      };
    }
  }

  // Adicionar último bloco
  blocos.push(blocoAtual);

  return blocos;
}

/**
 * Formata data para string legível (DD/MM/YYYY)
 */
function formatarData(data) {
  const dia = String(data.getDate()).padStart(2, '0');
  const mes = String(data.getMonth() + 1).padStart(2, '0');
  const ano = data.getFullYear();
  return dia + '/' + mes + '/' + ano;
}

/**
 * Obtém o calendário do utilizador
 */
function obterCalendario() {
  try {
    // Tentar obter calendário principal
    const calendarioPrincipal = CalendarApp.getDefaultCalendar();

    // Se o nome corresponder, usar este
    if (calendarioPrincipal.getName() === CONFIG.CALENDARIO.NOME) {
      return calendarioPrincipal;
    }

    // Procurar em todos os calendários próprios
    const calendarios = CalendarApp.getAllOwnedCalendars();
    for (let cal of calendarios) {
      if (cal.getName() === CONFIG.CALENDARIO.NOME) {
        return cal;
      }
    }

    // Se não encontrou calendário específico, usar o principal
    Logger.log('⚠️ Calendário "' + CONFIG.CALENDARIO.NOME + '" não encontrado. A usar calendário principal.');
    return calendarioPrincipal;

  } catch (erro) {
    Logger.log('❌ Erro ao obter calendário: ' + erro.message);
    return null;
  }
}

/**
 * Remove eventos "Férias" existentes no ano configurado
 * CORRIGIDO: Remove eventos que começam com "Férias" (inclui "Férias (X dias)")
 */
function limparEventosAntigos(calendario, ano) {
  const dataInicio = new Date(ano, 0, 1);  // 1 Janeiro
  const dataFim = new Date(ano + 1, 0, 1); // 1 Janeiro do ano seguinte (inclui 31 Dezembro)

  const eventosExistentes = calendario.getEvents(dataInicio, dataFim);

  let removidos = 0;
  eventosExistentes.forEach(evento => {
    const titulo = evento.getTitle();
    const descricao = evento.getDescription() || '';
    // Remover eventos criados pelo script (t?tulo "F?rias" ou marcador na descri??o)
    const geradoPeloScript =
      titulo.startsWith(CONFIG.CALENDARIO.TITULO_EVENTO) ||
      descricao.indexOf(CONFIG.CALENDARIO.MARCADOR) !== -1;

    if (geradoPeloScript) {
      evento.deleteEvent();
      removidos++;
      Logger.log('Removido: "' + titulo + '"');
    }
  });

  if (removidos > 0) {
    Logger.log('Total: ' + removidos + ' evento(s) antigo(s) removido(s) para ' + ano);
  } else {
    Logger.log('?? Nenhum evento antigo encontrado para remover em ' + ano);
  }
}

// ============================
// GEST?O DE TRIGGERS (AUTOMA??O)
// ============================
// GESTÃO DE TRIGGERS (AUTOMAÇÃO)
// ============================

/**
 * Instala trigger para atualizar automaticamente quando o sheet é editado
 * NOTA: Este trigger NÃO deteta mudanças de cores, apenas de valores
 */
function instalarTrigger() {
  try {
    // Remover triggers existentes da mesma função (evitar duplicação)
    removerTriggersExistentes();

    // Criar novo trigger onEdit (para mudanças de valores)
    ScriptApp.newTrigger('atualizarContadores')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onEdit()
      .create();

    Logger.log('✅ Trigger onEdit instalado');
    mostrarNotificacao('Atualização automática ativada!', 'Trigger instalado', 3);

  } catch (erro) {
    Logger.log('❌ Erro ao instalar trigger: ' + erro.message);
    mostrarNotificacao('Erro ao ativar atualização automática.', 'Erro', 5);
  }
}

/**
 * Instala trigger de tempo para sincronização automática a cada 5 minutos
 * Atualiza contadores E sincroniza com Calendar automaticamente
 */
function instalarTriggerAutomatico() {
  try {
    // Remover triggers de tempo existentes
    removerTriggersAutomaticos();

    // Criar trigger que executa a cada 5 minutos
    ScriptApp.newTrigger('sincronizarTudo')
      .timeBased()
      .everyMinutes(5)
      .create();

    Logger.log('✅ Trigger automático (5 min) instalado');
    mostrarNotificacao(
      'Sincronização automática ativada! Atualiza a cada 5 minutos.',
      'Automação Total Ativa',
      5
    );

  } catch (erro) {
    Logger.log('❌ Erro ao instalar trigger automático: ' + erro.message);
    mostrarNotificacao('Erro ao ativar sincronização automática.', 'Erro', 5);
  }
}

/**
 * Remove o trigger automático
 */
function removerTrigger() {
  try {
    removerTriggersExistentes();

    Logger.log('✅ Trigger onEdit removido');
    mostrarNotificacao('Atualização automática desativada!', 'Trigger removido', 3);

  } catch (erro) {
    Logger.log('❌ Erro ao remover trigger: ' + erro.message);
  }
}

/**
 * Remove o trigger de sincronização automática (5 minutos)
 */
function removerTriggerAutomatico() {
  try {
    removerTriggersAutomaticos();

    Logger.log('✅ Trigger automático (5 min) removido');
    mostrarNotificacao('Sincronização automática desativada!', 'Automação Total Desativada', 3);

  } catch (erro) {
    Logger.log('❌ Erro ao remover trigger automático: ' + erro.message);
  }
}

/**
 * Remove todos os triggers existentes da função atualizarContadores
 */
function removerTriggersExistentes() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'atualizarContadores') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * Remove todos os triggers de tempo (sincronizarTudo)
 */
function removerTriggersAutomaticos() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sincronizarTudo') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// ============================
// MENU PERSONALIZADO
// ============================

/**
 * Cria menu personalizado ao abrir o Google Sheet
 * Executado automaticamente pelo trigger onOpen
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🏖️ Gestão de Férias')
    .addItem('⚡ SINCRONIZAR TUDO', 'sincronizarTudo')
    .addSeparator()
    .addItem('🔄 Atualizar Contadores', 'atualizarContadores')
    .addItem('📅 Sincronizar com Calendar', 'sincronizarComCalendar')
    .addSeparator()
    .addItem('⚙️ Ativar Atualização ao Editar', 'instalarTrigger')
    .addItem('🤖 Ativar Sincronização Automática (5 min)', 'instalarTriggerAutomatico')
    .addSeparator()
    .addItem('❌ Desativar Atualização ao Editar', 'removerTrigger')
    .addItem('⛔ Desativar Sincronização Automática', 'removerTriggerAutomatico')
    .addSeparator()
    .addItem('🔍 Testar Deteção de Cores', 'testarDetecaoCores')
    .addItem('ℹ️ Ajuda', 'mostrarAjuda')
    .addToUi();
}

/**
 * Mostra janela de ajuda com instruções
 */
function mostrarAjuda() {
  const ui = SpreadsheetApp.getUi();

  const mensagem = [
    'GESTAO DE FERIAS ' + obterAnoDaSheet(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()),
    '',
    'SINCRONIZAR TUDO (RECOMENDADO)',
    '- Menu: "SINCRONIZAR TUDO"',
    '- Atualiza contadores e sincroniza Calendar',
    '- Usa sempre que pintares celulas de ferias',
    '',
    'SINCRONIZACAO AUTOMATICA',
    '- Menu: "Ativar Sincronizacao Automatica (5 min)"',
    '- Atualiza tudo automaticamente a cada 5 minutos',
    '- Pinta celulas e esquece - o sistema faz o resto',
    '',
    'CONTADORES MANUAIS',
    '- Menu: "Atualizar Contadores" - so numeros',
    '- Menu: "Sincronizar com Calendar" - so eventos',
    '',
    'CORES A USAR',
    '- Roxo (#d9d2e9): Ferias planeadas',
    '- Verde (#d9ead3): Dia de aniversario da empresa',
    '',
    'RECOMENDACOES',
    'OPCAO 1 - Totalmente automatico:',
    '1. Ativa "Sincronizacao Automatica (5 min)"',
    '2. Pinta celulas a vontade',
    '3. Aguarda ate 5 minutos',
    '4. Tudo atualiza sozinho',
    '',
    'OPCAO 2 - Semi-automatico:',
    '1. Pinta as celulas de ferias',
    '2. Clica em "SINCRONIZAR TUDO"',
    '3. Pronto! (instantaneo)',
    '',
    'Se houver problemas, usa "Testar Detecao de Cores"',
  ].join('\n');

  ui.alert('Ajuda - Gestao de Ferias', mensagem, ui.ButtonSet.OK);
}

// ============================
// CONFIGURAÇÃO INICIAL
// ============================

/**
 * Configura o Google Sheet pela primeira vez
 * Cria a estrutura de legenda e contadores
 * EXECUTAR UMA VEZ após colar o código
 */
function configurarSheet() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Estrutura da legenda (a partir da linha 18, coluna B)
    const legenda = [
      ['Férias ano corrente disponíveis', 0],
      ['Férias transitadas do ano anterior', 0],
      ['Férias gozadas', 0],
      ['Férias planeadas', 0],
      ['Total (gozadas + planeadas)', 0],
      ['Férias restantes', 0],
      ['', ''],
      ['Dia de aniversário disponível', 1],
      ['Dia de aniversário gozado', 0],
      ['Dia de aniversário a gozar', 0],
      ['Dia de aniversário restante', 0]
    ];

    // Inserir dados na linha 18, coluna B
    sheet.getRange(18, 2, legenda.length, 2).setValues(legenda);

    // Aplicar formatação às células da legenda
    sheet.getRange('B18:B23').setFontWeight('normal');
    sheet.getRange('B18').setFontWeight('bold'); // destaque primeira linha
    sheet.getRange('B24').setFontWeight('bold'); // separador antes do aniversário

    // Aplicar cores às células da legenda (exemplos visuais)
    sheet.getRange('B18:C18').setBackground('#e6b8af'); // disponíveis corrente
    sheet.getRange('B19:C19').setBackground('#fff2cc'); // dias transitados ano anterior
    sheet.getRange('B20:C20').setBackground('#d9d2e9'); // gozadas
    sheet.getRange('B21:C21').setBackground('#d9d2e9'); // planeadas
    sheet.getRange('B22:C22').setBackground('#d9d2e9'); // total
    sheet.getRange('B23:C23').setBackground('#b6d7a8'); // restantes
    sheet.getRange('B25:C27').setBackground('#d9ead3'); // aniversário

    Logger.log('✅ Sheet configurado com sucesso!');
    mostrarNotificacao('Legenda e contadores configurados!', 'Configuração completa', 3);

    // Executar atualização inicial dos contadores
    atualizarContadores(null, sheet, obterAnoDaSheet(sheet));

  } catch (erro) {
    Logger.log('❌ Erro ao configurar sheet: ' + erro.message);
    mostrarNotificacao('Erro na configuração. Verifica o log.', 'Erro', 5);
  }
}

// ============================
// FUNÇÕES AUXILIARES
// ============================

/**
 * Mostra notificação toast no Google Sheets
 */
function mostrarNotificacao(mensagem, titulo, duracao) {
  SpreadsheetApp.getActiveSpreadsheet().toast(mensagem, titulo, duracao);
}

/**
 * Função de teste para validar cores
 * Executa e verifica o log para ver as cores detetadas
 * 
 * COMO USAR:
 * 1. No Apps Script Editor, seleciona esta função no dropdown
 * 2. Clica em Executar (▶️)
 * 3. Vai a "Execuções" (ou View > Logs)
 * 4. Vê todas as cores encontradas no teu calendário
 */
function testarDetecaoCores() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(CONFIG.CALENDAR_RANGE);
  const valores = range.getValues();
  const cores = range.getBackgrounds();

  const coresEncontradas = new Map(); // Guarda cor e quantidade de ocorrências

  // Percorrer todas as células
  for (let i = 0; i < cores.length; i++) {
    for (let j = 0; j < cores[i].length; j++) {
      const cor = cores[i][j];
      const valor = valores[i][j];

      // Ignorar branco e preto, e células vazias ou com texto
      if (cor && cor !== '#ffffff' && cor !== '#000000' && valor && !isNaN(valor)) {
        const corNormalizada = cor.toLowerCase().replace(/\s/g, '');

        if (coresEncontradas.has(corNormalizada)) {
          coresEncontradas.set(corNormalizada, coresEncontradas.get(corNormalizada) + 1);
        } else {
          coresEncontradas.set(corNormalizada, 1);
        }
      }
    }
  }

  // Mostrar resultados no log
  Logger.log('═══════════════════════════════════════');
  Logger.log('🎨 DIAGNÓSTICO DE CORES - CALENDÁRIO ' + obterAnoDaSheet(sheet));
  Logger.log('═══════════════════════════════════════\n');

  Logger.log('📊 Cores encontradas no calendário (apenas células com números):');
  coresEncontradas.forEach((quantidade, cor) => {
    Logger.log('  ' + cor + ' : ' + quantidade + ' celulas');
  });

  Logger.log('\n🎯 Cores configuradas para deteção:');
  Logger.log('  Ferias (ano corrente): ' + CONFIG.CORES.FERIAS_ATUAL.toLowerCase());
  Logger.log('  Ferias (ano corrente - alt): ' + CONFIG.CORES.FERIAS_ATUAL_ALT.toLowerCase());
  Logger.log('  Ferias (ano anterior): ' + CONFIG.CORES.FERIAS_ANTERIOR.toLowerCase());
  Logger.log('  Aniversario (verde): ' + CONFIG.CORES.ANIVERSARIO.toLowerCase());

  Logger.log('\n✅ Correspondências encontradas:');
  let encontrouFerias = false;
  let encontrouAniversario = false;

  coresEncontradas.forEach((quantidade, cor) => {
    if (cor === CONFIG.CORES.FERIAS_ATUAL.toLowerCase()) {
      Logger.log('Ferias (ano corrente): ' + quantidade + ' celulas detetadas!');
      encontrouFerias = true;
    }
    if (cor === CONFIG.CORES.FERIAS_ATUAL_ALT.toLowerCase()) {
      Logger.log('Ferias (ano corrente - alt): ' + quantidade + ' celulas detetadas!');
      encontrouFerias = true;
    }
    if (cor === CONFIG.CORES.FERIAS_ANTERIOR.toLowerCase()) {
      Logger.log('Ferias (ano anterior): ' + quantidade + ' celulas detetadas!');
      encontrouFerias = true;
    }
    if (cor === CONFIG.CORES.ANIVERSARIO.toLowerCase()) {
      Logger.log('Aniversario: ' + quantidade + ' celulas verdes detetadas!');
      encontrouAniversario = true;
    }
  });

  if (!encontrouFerias) {
    Logger.log('  ✗ Nenhuma célula roxa de férias encontrada!');
    Logger.log('    → Verifica se usaste a cor: ' + CONFIG.CORES.FERIAS);
  }

  if (!encontrouAniversario) {
    Logger.log('  ✗ Nenhuma célula verde de aniversário encontrada!');
    Logger.log('    → Verifica se usaste a cor: ' + CONFIG.CORES.ANIVERSARIO);
  }

  Logger.log('\n═══════════════════════════════════════');

  // Mostrar também como notificação
  let mensagem = 'Cores encontradas: ' + coresEncontradas.size + '\n';
  mensagem += encontrouFerias ? '✓ Férias OK\n' : '✗ Férias não encontradas\n';
  mensagem += encontrouAniversario ? '✓ Aniversário OK' : '✗ Aniversário não encontrado';

  mostrarNotificacao(mensagem, 'Teste de Cores', 8);
}

/**
 * Atualiza as cores no CONFIG baseado no que está pintado no sheet
 * EXECUTAR DEPOIS DE testarDetecaoCores() se as cores não corresponderem
 */
function atualizarCoresAutomaticamente() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(CONFIG.CALENDAR_RANGE);
  const valores = range.getValues();
  const cores = range.getBackgrounds();

  const coresEncontradas = new Map();

  // Contar ocorrências de cada cor
  for (let i = 0; i < cores.length; i++) {
    for (let j = 0; j < cores[i].length; j++) {
      const cor = cores[i][j];
      const valor = valores[i][j];

      if (cor && cor !== '#ffffff' && cor !== '#000000' && valor && !isNaN(valor)) {
        const corNormalizada = cor.toLowerCase().replace(/\s/g, '');
        coresEncontradas.set(corNormalizada, (coresEncontradas.get(corNormalizada) || 0) + 1);
      }
    }
  }

  // Ordenar por quantidade (mais usadas primeiro)
  const coresOrdenadas = Array.from(coresEncontradas.entries()).sort((a, b) => b[1] - a[1]);

  if (coresOrdenadas.length >= 1) {
    const corMaisUsada = coresOrdenadas[0][0];
    Logger.log('Sugestao: Definir cor de ferias como: ' + corMaisUsada);
    Logger.log('  (Encontradas ' + coresOrdenadas[0][1] + ' celulas com esta cor)');
  }

  if (coresOrdenadas.length >= 2) {
    const segundaCorMaisUsada = coresOrdenadas[1][0];
    Logger.log('Sugestao: Definir cor de aniversario como: ' + segundaCorMaisUsada);
    Logger.log('  (Encontradas ' + coresOrdenadas[1][1] + ' celulas com esta cor)');
  }

  Logger.log('\n💡 Para atualizar as cores no código:');
  Logger.log('1. Edita o objeto CONFIG no topo do código');
  Logger.log('2. Altera os valores em CONFIG.CORES.FERIAS e CONFIG.CORES.ANIVERSARIO');
  Logger.log('3. Guarda o script e executa atualizarContadores() novamente');
}
