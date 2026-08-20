/**
 * SISTEMA DE GESTÃO DE FÉRIAS
 * Versão: 1.5.1
 * Data: 2026-08-20
 *
 * Autor: Emanuel Ferreira (@emanuwells)
 *
 * Descrição:
 * Sistema completo para gestão de férias com:
 * - Contagem automática de dias de férias (células roxas).
 * - Contagem automática do dia de aniversário (células verdes).
 * - Sincronização com Google Calendar por diferença (cria/atualiza/remove só o que mudou),
 *   com agrupamento de dias consecutivos, debounce ao pintar e recuperação automática de
 *   quota esgotada do Calendar (mesmo padrão do projeto irmão Luna Sheet).
 * - Cada bloco de férias sincroniza de forma isolada: um evento com problema fica
 *   registado como falha, sem abortar a sincronização dos restantes blocos.
 * - Atualização automática via trigger diário + onChange.
 * - Menu personalizado.
 *
 */

// ============================
// CONFIGURAÇÕES GLOBAIS
// ============================

const CONFIG = {
  // Intervalo do calendário: 12 linhas (meses) por 31 colunas (dias).
  // Ajusta aqui se mudares a posição do quadro; no teu layout, o topo do calendário começa em C5 e o dia 31 cai em AM16.
  CALENDAR_RANGE: 'C5:AM16',

  // Cores a detetar, em hexadecimal, no formato do Google Sheets.
  CORES: {
    FERIAS_ATUAL: '#d9d2e9',      // Roxo/lilás: férias planeadas do ano corrente.
    FERIAS_ATUAL_ALT: '#b4a7d6',  // Variante comum de roxo/lilás nas folhas.
    FERIAS_ANTERIOR: '#fff2cc',   // Amarelo claro: dias transitados do ano anterior.
    ANIVERSARIO: '#d9ead3'        // Verde claro: dia de aniversário.
  },

  // Células onde aparecem os contadores, na coluna C.
  CELULAS: {
    // Contadores de férias.
    FERIAS_DISPONIVEIS: 'C18',      // Input manual do utilizador, para o ano corrente.
    FERIAS_ANTERIOR: 'C19',         // Dias transitados do ano anterior, com input manual.
    FERIAS_GOZADAS: 'C20',          // Calculado automaticamente.
    FERIAS_PLANEADAS: 'C21',        // Calculado automaticamente.
    FERIAS_TOTAL: 'C22',            // Soma de gozadas e planeadas.
    FERIAS_RESTANTES: 'C23',        // Diferença entre disponíveis, anteriores e total.

    // Contadores de aniversário.
    ANIVERSARIO_DISPONIVEL: 'C25', // Fixo: 1 dia.
    ANIVERSARIO_GOZADO: 'C26',     // 0 ou 1, se a data já passou.
    ANIVERSARIO_A_GOZAR: 'C27'     // 0 ou 1, se a data for futura.
  },

  // Configurações do Google Calendar.
  CALENDARIO: {
    NOME: '',                      // Deixe vazio para usar o calendário principal (recomendado).
    TITULO_EVENTO: 'Férias',       // Título dos eventos criados.
    ANO: new Date().getFullYear(), // Ano por omissão, substituído pelo ano da folha se existir.
    MARCADOR: '[FERIAS_AUTO]'      // Assinatura para identificar eventos gerados pelo script.
  },

  // Sincronização com o Calendar: debounce ao pintar e recuperação de quota (mesmo padrão do Luna Sheet).
  SYNC: {
    EDIT_SYNC_DELAY_MS: 5 * 60 * 1000,        // Atraso do Calendar após pintar/editar (contadores continuam imediatos).
    QUOTA_RETRY_DELAY_MS: 6 * 60 * 60 * 1000, // Atraso da retentativa automática após esgotar a quota diária do Calendar.
    MAX_EVENT_MUTATIONS_PER_RUN: 40           // Limite defensivo de criações/atualizações/remoções por execução.
  }
};

/** Chave usada para ignorar onChange provocado por escrita do próprio script. */
const CHAVE_SUPRESSAO_ONCHANGE = 'VACATION_MODE_SUPPRESS_ONCHANGE';

/** Chave de PropertiesService (script-wide) para o bloqueio de quota do Calendar. */
const PROPRIEDADES = {
  quotaRetryAt: 'VACATION_MODE_CALENDAR_QUOTA_RETRY_AT'
};

/** Nomes de funções usados pelos triggers, para instalar/remover de forma consistente. */
const HANDLERS = {
  sincronizacaoDiaria: 'sincronizarTudo',
  alteracaoFolha: 'onAlteracaoPlanilha',
  calendarPendente: 'sincronizarCalendarPendente'
};

/** Mensagem exata que o Google devolve quando a quota diária do Calendar esgota. */
const REGEX_QUOTA_CALENDAR = /Service invoked too many times(?: for one day)?:\s*calendar/i;

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
 * Devolve as folhas de calendário anuais encontradas no ficheiro.
 * Aceita nomes como "Calendário 2026", "Calendário de férias" ou folhas com ano no nome.
 * Se nenhuma for encontrada, devolve apenas a folha ativa como fallback.
 */
function obterFolhasCalendario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const todas = ss.getSheets();

  const comAnoNoNome = todas
    .map(sheet => ({ sheet, ano: obterAnoDaSheet(sheet) }))
    .filter(({ sheet }) => /Calend.?rio/i.test(sheet.getName()) && /20\d{2}/.test(sheet.getName()));

  if (comAnoNoNome.length > 0) {
    return comAnoNoNome;
  }

  const comNomeCalendario = todas
    .map(sheet => ({ sheet, ano: obterAnoDaSheet(sheet) }))
    .filter(({ sheet }) => /Calend.?rio|f[eé]rias/i.test(sheet.getName()));

  if (comNomeCalendario.length > 0) {
    return comNomeCalendario;
  }

  if (todas.length === 1) {
    const unica = todas[0];
    return [{ sheet: unica, ano: obterAnoDaSheet(unica) }];
  }

  const ativa = ss.getActiveSheet();
  return [{ sheet: ativa, ano: obterAnoDaSheet(ativa) }];
}

/**
 * Executa uma ação sem voltar a disparar sincronização por onChange.
 */
function comSupressaoAlteracaoFolha(callback) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty(CHAVE_SUPRESSAO_ONCHANGE, '1');
  try {
    return callback();
  } finally {
    props.deleteProperty(CHAVE_SUPRESSAO_ONCHANGE);
  }
}

/**
 * Indica se a sincronização automática deve ser ignorada neste momento.
 */
function deveIgnorarOnChange() {
  return PropertiesService.getDocumentProperties().getProperty(CHAVE_SUPRESSAO_ONCHANGE) === '1';
}

// ============================
// ============================
// FUNÇÃO PRINCIPAL - ATUALIZAR CONTADORES
// ============================

/**
 * Atualiza todos os contadores de férias e aniversário.
 * Conta células coloridas e distingue entre datas passadas e futuras.
 */
function atualizarContadores(e, sheetParam, anoParam, silencioso) {
  try {
    const sheet = sheetParam || (e && e.range ? e.range.getSheet() : SpreadsheetApp.getActiveSpreadsheet().getActiveSheet());
    const ano = anoParam || obterAnoDaSheet(sheet);
    const hoje = obterDataHoje();

    // Obter dados do calendário: valores e cores de fundo.
    const range = sheet.getRange(CONFIG.CALENDAR_RANGE);
    const valores = range.getValues();
    const cores = range.getBackgrounds();

    // Inicializar contadores.
    const contadores = {
      feriasGozadas: 0,
      feriasPlaneadas: 0,
      aniversarioGozado: 0,
      aniversarioAGozar: 0
    };

    // Percorrer todas as células do calendário.
    for (let linha = 0; linha < valores.length; linha++) {
      for (let coluna = 0; coluna < valores[linha].length; coluna++) {
        processarCelula(valores[linha][coluna], cores[linha][coluna], linha, hoje, contadores, ano);
      }
    }

    // Atualizar células na folha com os novos valores, sem reentrar no onChange.
    comSupressaoAlteracaoFolha(function () {
      atualizarCelulasContadores(sheet, contadores);
    });

    Logger.log('Contadores atualizados com sucesso (' + sheet.getName() + ' - ' + ano + ')');
    if (!silencioso) {
      mostrarNotificacao(sheet.getName() + ': Contadores atualizados!', 'Sucesso', 3);
    }

  } catch (erro) {
    Logger.log('Erro ao atualizar contadores: ' + erro.message);
    if (!silencioso) {
      mostrarNotificacao('Erro ao atualizar contadores. Verifica o log.', 'Erro', 5);
    }
  }
}

/**
 * Obtém a data de hoje normalizada, sem horas.
 */
function obterDataHoje() {
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  return hoje;
}

/**
 * Normaliza uma cor para comparação estável com os valores do Google Sheets.
 */
function normalizarCor(cor) {
  return String(cor || '').toLowerCase().replace(/\s/g, '');
}

/**
 * Indica se a cor corresponde a uma das cores configuradas para férias.
 */
function isCorFerias(cor) {
  const corNormalizada = normalizarCor(cor);
  return [
    CONFIG.CORES.FERIAS_ATUAL,
    CONFIG.CORES.FERIAS_ATUAL_ALT,
    CONFIG.CORES.FERIAS_ANTERIOR
  ].some(corConfigurada => corNormalizada === corConfigurada.toLowerCase());
}

/**
 * Indica se a cor corresponde à cor configurada para o dia de aniversário.
 */
function isCorAniversario(cor) {
  return normalizarCor(cor) === CONFIG.CORES.ANIVERSARIO.toLowerCase();
}

/**
 * Extrai o dia do mês a partir de números, texto numérico ou datas reais.
 */
function extrairDiaDaCelula(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return valor.getDate();
  }

  if (typeof valor === 'number' && isFinite(valor)) {
    return Math.trunc(valor);
  }

  const match = String(valor || '').trim().match(/^(\d{1,2})$/);
  return match ? parseInt(match[1], 10) : null;
}

/**
 * Converte uma célula da grelha anual numa data válida do ano indicado.
 */
function obterDataDaCelula(valor, indiceLinha, ano) {
  const dia = extrairDiaDaCelula(valor);
  if (!dia) {
    return null;
  }

  const mes = indiceLinha; // 0=Janeiro, 1=Fevereiro, ..., 11=Dezembro.
  const data = new Date(ano, mes, dia);

  if (data.getFullYear() !== ano || data.getMonth() !== mes || data.getDate() !== dia) {
    return null;
  }

  data.setHours(0, 0, 0, 0);
  return data;
}

/**
 * Processa uma célula individual e atualiza os contadores.
 */
function processarCelula(valor, cor, indiceLinha, hoje, contadores, ano) {
  // Ignorar células vazias, texto de dias da semana e datas inválidas.
  const data = obterDataDaCelula(valor, indiceLinha, ano);
  if (!data) {
    return;
  }

  if (isCorFerias(cor)) {
    if (data <= hoje) {
      contadores.feriasGozadas++;
    } else {
      contadores.feriasPlaneadas++;
    }
  }

  // Verificar se é célula de aniversário (verde).
  if (isCorAniversario(cor)) {
    if (data <= hoje) {
      contadores.aniversarioGozado = 1; // Máximo: 1 dia.
    } else {
      contadores.aniversarioAGozar = 1; // Máximo: 1 dia.
    }
  }
}

/**
 * Escreve os contadores calculados nas células configuradas da folha.
 */
function atualizarCelulasContadores(sheet, contadores) {
  // Atualizar contadores de férias.
  sheet.getRange(CONFIG.CELULAS.FERIAS_GOZADAS).setValue(contadores.feriasGozadas);
  sheet.getRange(CONFIG.CELULAS.FERIAS_PLANEADAS).setValue(contadores.feriasPlaneadas);

  const totalPlaneado = contadores.feriasGozadas + contadores.feriasPlaneadas;
  sheet.getRange(CONFIG.CELULAS.FERIAS_TOTAL).setValue(totalPlaneado);

  // Calcular e atualizar dias restantes.
  const disponiveis = sheet.getRange(CONFIG.CELULAS.FERIAS_DISPONIVEIS).getValue() || 0;
  const anterior = sheet.getRange(CONFIG.CELULAS.FERIAS_ANTERIOR).getValue() || 0;
  const restantes = (disponiveis + anterior) - totalPlaneado;
  sheet.getRange(CONFIG.CELULAS.FERIAS_RESTANTES).setValue(restantes);

  // Atualizar contadores de aniversário.
  sheet.getRange(CONFIG.CELULAS.ANIVERSARIO_GOZADO).setValue(contadores.aniversarioGozado);
  sheet.getRange(CONFIG.CELULAS.ANIVERSARIO_A_GOZAR).setValue(contadores.aniversarioAGozar);
}

// ============================
// SINCRONIZAÇÃO COMPLETA (CONTADORES + CALENDAR)
// ============================

/**
 * Sincroniza tudo: atualiza contadores e sincroniza com o Google Calendar.
 * Função combinada para usar com botão ou trigger automático.
 */
function sincronizarTudo(opcoes) {
  const automatico = opcoes && opcoes.automatico === true;
  try {
    Logger.log('A iniciar sincronização completa...');

    const folhas = obterFolhasCalendario();
    if (folhas.length === 0) {
      Logger.log('Nenhuma folha de calendário encontrada.');
      if (!automatico) {
        mostrarNotificacao('Nenhuma folha de calendário encontrada.', 'Aviso', 5);
      }
      return;
    }

    folhas.forEach(({ sheet, ano }) => {
      atualizarContadores(null, sheet, ano, automatico);
      Utilities.sleep(500);
      sincronizarComCalendar(sheet, ano, automatico);
    });

    Logger.log('Sincronização completa finalizada!');
    if (!automatico) {
      mostrarNotificacao('Contadores e Calendar sincronizados!', 'Sincronização Completa', 5);
    }

  } catch (erro) {
    Logger.log('Erro na sincronização completa: ' + erro.message);
    if (!automatico) {
      mostrarNotificacao('Erro na sincronização. Verifica o log.', 'Erro', 5);
    }
  }
}

/**
 * Handler de onChange para sincronizar ao colorir ou editar a grelha anual.
 * Ignora alterações feitas pelo próprio script e tipos de mudança irrelevantes.
 * Os contadores atualizam de imediato (só leem/escrevem a folha); o Calendar é
 * agendado com debounce, para não gastar a quota diária a cada pintura.
 */
function onAlteracaoPlanilha(e) {
  if (!e || deveIgnorarOnChange()) {
    return;
  }

  const tiposRelevantes = { FORMAT: true, EDIT: true };
  if (!tiposRelevantes[e.changeType]) {
    return;
  }

  Logger.log('Alteração detetada (' + e.changeType + '). A atualizar contadores e a agendar sincronização do Calendar...');

  obterFolhasCalendario().forEach(({ sheet, ano }) => {
    atualizarContadores(null, sheet, ano, true);
  });

  agendarSincronizacaoCalendar();
}

/**
 * Agenda uma sincronização única do Calendar, adiada, para agregar pinturas sucessivas
 * num único ciclo em vez de disparar uma sincronização completa a cada alteração.
 * Se a quota diária estiver bloqueada, não agenda nada (a retentativa já está prevista).
 */
function agendarSincronizacaoCalendar() {
  const retryAt = Number(PropertiesService.getScriptProperties().getProperty(PROPRIEDADES.quotaRetryAt) || 0);
  if (retryAt > Date.now()) {
    Logger.log('Sincronização do Calendar adiada devido à quota até ' + new Date(retryAt).toISOString() + '.');
    return;
  }

  agendarTriggerUnico(HANDLERS.calendarPendente, CONFIG.SYNC.EDIT_SYNC_DELAY_MS);
}

/**
 * Cria um trigger único (baseado em tempo) para o handler indicado, removendo qualquer
 * trigger pendente do mesmo handler primeiro. Usa lock para evitar duplicados quando
 * várias alterações disparam o agendamento quase em simultâneo.
 */
function agendarTriggerUnico(handler, atrasoMs) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log('Agendamento ignorado: outra execução está a atualizar os triggers.');
    return false;
  }

  try {
    removerTriggersPorHandler(handler);
    ScriptApp.newTrigger(handler).timeBased().after(atrasoMs).create();
    return true;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Remove todos os triggers do projeto associados a um handler específico.
 */
function removerTriggersPorHandler(handler) {
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === handler)
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

/**
 * Executa a sincronização do Calendar pendente (fim do debounce ou retentativa de quota)
 * para todas as folhas de calendário detetadas.
 */
function sincronizarCalendarPendente() {
  removerTriggersPorHandler(HANDLERS.calendarPendente);
  PropertiesService.getScriptProperties().deleteProperty(PROPRIEDADES.quotaRetryAt);

  const folhas = obterFolhasCalendario();
  folhas.forEach(({ sheet, ano }) => {
    sincronizarComCalendar(sheet, ano, true);
  });
}

/**
 * Sincroniza uma folha específica com o Calendar (usa o ano detetado na folha).
 * Sincronização por diferença: só cria, atualiza ou apaga o que realmente mudou desde
 * a última execução (em vez de apagar e recriar o ano inteiro). Respeita um bloqueio
 * local de quota diária do Calendar; execuções manuais limpam esse bloqueio e forçam
 * uma tentativa real, execuções automáticas saem em silêncio até à retentativa agendada.
 */
function sincronizarComCalendar(sheetParam, anoParam, silencioso) {
  let lock;
  const props = PropertiesService.getScriptProperties();
  try {
    const sheet = sheetParam || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const ano = anoParam || obterAnoDaSheet(sheet);

    const retryAt = Number(props.getProperty(PROPRIEDADES.quotaRetryAt) || 0);
    if (retryAt > Date.now()) {
      if (silencioso) {
        Logger.log('Sincronização do Calendar suspensa até ' + new Date(retryAt).toISOString() + ' devido à quota diária.');
        return;
      }
      props.deleteProperty(PROPRIEDADES.quotaRetryAt);
      Logger.log('Execução manual: bloqueio local de quota removido; o Calendar será testado agora.');
    }

    lock = LockService.getDocumentLock();
    if (!lock.tryLock(30000)) {
      Logger.log('Outra sincronização está a correr. Abortado para evitar duplicados.');
      if (!silencioso) {
        mostrarNotificacao('Outra sincronização em curso. Tenta novamente em instantes.', 'Aviso', 4);
      }
      return;
    }
    Logger.log('A iniciar sincronização com Google Calendar para ' + sheet.getName() + ' (' + ano + ')...');

    // Obter ou aceder ao calendário.
    const calendario = obterCalendario();
    if (!calendario) {
      Logger.log('Erro: calendário não encontrado');
      if (!silencioso) {
        mostrarNotificacao('Erro ao aceder ao calendário. Verifica as permissões.', 'Erro', 5);
      }
      return;
    }

    Logger.log('Calendário obtido: ' + calendario.getName());

    // Obter todas as datas de férias e agrupar em blocos (dias úteis seguidos, com fins de semana contíguos).
    const datasFerias = obterDatasFerias(sheet, ano);
    const feriasRestantes = sheet.getRange(CONFIG.CELULAS.FERIAS_RESTANTES).getValue() || 0;
    const blocos = agruparDatasConsecutivas(datasFerias);
    const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

    Logger.log('Total de dias de férias encontrados (' + sheet.getName() + '): ' + datasFerias.length +
      ', agrupados em ' + blocos.length + ' período(s).');

    const eventosDesejados = blocos.map(bloco => construirEventoDesejado(bloco, feriasRestantes, sheetUrl));
    const eventosExistentes = obterEventosGeradosNoAno(calendario, ano);

    const resultado = sincronizarBlocosComDiferenca(calendario, eventosDesejados, eventosExistentes);

    Logger.log('Sincronização concluída (' + sheet.getName() + '): ' + resultado.criados + ' criado(s), ' +
      resultado.atualizados + ' atualizado(s), ' + resultado.removidos + ' removido(s), ' +
      resultado.falhados + ' falhado(s).');

    if (resultado.pendente) {
      const agendada = agendarTriggerUnico(HANDLERS.calendarPendente, CONFIG.SYNC.EDIT_SYNC_DELAY_MS);
      Logger.log('Limite de alterações por execução atingido; continuação ' + (agendada ? 'agendada' : 'não agendada') + '.');
    }

    if (!silencioso) {
      if (resultado.criados === 0 && resultado.atualizados === 0 && resultado.removidos === 0 && resultado.falhados === 0) {
        mostrarNotificacao(sheet.getName() + ': Calendar já estava atualizado (sem alterações).', 'Sincronização completa', 3);
      } else {
        const partes = [];
        if (resultado.criados) partes.push(resultado.criados + ' criado(s)');
        if (resultado.atualizados) partes.push(resultado.atualizados + ' atualizado(s)');
        if (resultado.removidos) partes.push(resultado.removidos + ' removido(s)');
        if (resultado.falhados) partes.push(resultado.falhados + ' falhado(s) — ver log');
        mostrarNotificacao(sheet.getName() + ': ' + partes.join(', ') + ' no Google Calendar.', 'Sincronização completa', 5);
      }
    }

  } catch (erro) {
    const mensagem = erro && erro.message ? erro.message : String(erro);
    if (REGEX_QUOTA_CALENDAR.test(mensagem)) {
      const proxima = new Date(Date.now() + CONFIG.SYNC.QUOTA_RETRY_DELAY_MS);
      props.setProperty(PROPRIEDADES.quotaRetryAt, String(proxima.getTime()));
      const agendada = agendarTriggerUnico(HANDLERS.calendarPendente, CONFIG.SYNC.QUOTA_RETRY_DELAY_MS);
      Logger.log('Quota diária do Google Calendar esgotada. Retentativa ' + (agendada ? 'agendada' : 'não agendada') +
        ' para ' + proxima.toISOString() + '.');
      if (!silencioso) {
        mostrarNotificacao(
          'Quota diária do Google Calendar esgotada. Nova tentativa automática mais tarde, ou usa "Sincronizar com Calendar" para forçar agora.',
          'Quota do Calendar',
          6
        );
      }
    } else {
      Logger.log('Erro ao sincronizar com Calendar: ' + mensagem);
      Logger.log('Stack trace: ' + (erro && erro.stack));
      if (!silencioso) {
        mostrarNotificacao('Erro ao sincronizar. Verifica o log.', 'Erro', 5);
      }
    }
  } finally {
    if (lock) {
      lock.releaseLock();
    }
  }
}

/**
 * Constrói o evento "desejado" para um bloco de dias de férias: datas (com fins de semana
 * contíguos incluídos), título e descrição. A chave lógica é a data de início já estendida
 * (a mesma que fica gravada como início real do evento), o que mantém a correspondência
 * estável quando o bloco cresce para a frente e só cria um evento novo quando o primeiro
 * dia do período realmente muda.
 */
function construirEventoDesejado(bloco, feriasRestantes, sheetUrl) {
  const intervalo = estenderIntervaloComFinsDeSemanaContiguos(bloco.inicio, bloco.fim);
  const dataInicio = intervalo.inicio;
  const dataFimApi = new Date(intervalo.fim);
  dataFimApi.setDate(dataFimApi.getDate() + 1); // A Calendar API precisa do dia seguinte para eventos de dia inteiro.

  const diasUteis = bloco.dias.length;
  const titulo = diasUteis === 1
    ? CONFIG.CALENDARIO.TITULO_EVENTO
    : CONFIG.CALENDARIO.TITULO_EVENTO + ' (' + diasUteis + ' dias)';

  const resumoPeriodo = '📅 ' + diasUteis + ' dias úteis · ' + formatarData(dataInicio) + ' a ' + formatarData(intervalo.fim);
  const resumoRestantes = '📉 Restantes: ' + feriasRestantes + ' dias';
  const linkSheet = '🔗 Sheet: ' + sheetUrl;
  const descricao = [resumoPeriodo, resumoRestantes, linkSheet, CONFIG.CALENDARIO.MARCADOR].join('\n');

  return { chave: String(dataInicio.getTime()), inicio: dataInicio, fimApi: dataFimApi, titulo, descricao };
}

/**
 * Recolhe todas as datas pintadas com cores de férias na grelha anual.
 */
function obterDatasFerias(sheet, ano) {
  const range = sheet.getRange(CONFIG.CALENDAR_RANGE);
  const valores = range.getValues();
  const cores = range.getBackgrounds();
  const startRow = range.getRow();      // Linha real da primeira célula do calendário.
  const startCol = range.getColumn();   // Coluna real da primeira célula do calendário.

  const datas = [];

  // Percorrer todas as células.
  for (let linha = 0; linha < valores.length; linha++) {
    for (let coluna = 0; coluna < valores[linha].length; coluna++) {
      const valor = valores[linha][coluna];
      const cor = cores[linha][coluna];

      const data = obterDataDaCelula(valor, linha, ano);
      if (!data) {
        continue;
      }

      if (isCorFerias(cor)) {
        datas.push(data);
        Logger.log('Encontrada: ' + formatarData(data) + ' (linha ' + (startRow + linha) + ', coluna ' + (startCol + coluna) + ')');
      }
    }
  }

  // Ordenar datas cronologicamente.
  datas.sort((a, b) => a - b);

  return datas;
}

/**
 * Indica se todos os dias entre duas datas (exclusivo) são sábado ou domingo.
 */
function intervaloApenasFinsDeSemana(dataInicio, dataFim) {
  const diaAtual = new Date(dataInicio);
  diaAtual.setHours(0, 0, 0, 0);
  diaAtual.setDate(diaAtual.getDate() + 1);

  const limite = new Date(dataFim);
  limite.setHours(0, 0, 0, 0);

  while (diaAtual < limite) {
    const diaSemana = diaAtual.getDay();
    if (diaSemana !== 0 && diaSemana !== 6) {
      return false;
    }
    diaAtual.setDate(diaAtual.getDate() + 1);
  }

  return true;
}

/**
 * Conta dias de calendário entre duas datas, inclusive.
 */
function contarDiasCalendario(inicio, fim) {
  const diferenca = Math.round((fim - inicio) / (1000 * 60 * 60 * 24));
  return diferenca + 1;
}

/**
 * Alarga o intervalo do evento para incluir fins de semana contíguos nas extremidades.
 * Segunda-feira no início inclui o sábado e domingo anteriores; sexta-feira no fim inclui o fim de semana seguinte.
 */
function estenderIntervaloComFinsDeSemanaContiguos(inicio, fim) {
  const inicioExt = new Date(inicio);
  inicioExt.setHours(0, 0, 0, 0);
  const fimExt = new Date(fim);
  fimExt.setHours(0, 0, 0, 0);

  if (inicioExt.getDay() === 1) {
    inicioExt.setDate(inicioExt.getDate() - 2);
  }
  if (fimExt.getDay() === 5) {
    fimExt.setDate(fimExt.getDate() + 2);
  }

  return { inicio: inicioExt, fim: fimExt };
}

/**
 * Agrupa datas consecutivas em blocos para criar eventos de dia inteiro.
 * Dias úteis separados apenas por fins de semana não pintados são fundidos num único bloco.
 */
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

    const diferencaDias = Math.round((dataAtual - dataAnterior) / (1000 * 60 * 60 * 24));
    const consecutivo = diferencaDias === 1;
    const separadoApenasPorFimDeSemana = diferencaDias > 1 && intervaloApenasFinsDeSemana(dataAnterior, dataAtual);

    Logger.log('Comparando ' + formatarData(dataAnterior) + ' com ' + formatarData(dataAtual) + ': diferenca = ' + diferencaDias + ' dia(s)');

    if (consecutivo || separadoApenasPorFimDeSemana) {
      blocoAtual.fim = dataAtual;
      blocoAtual.dias.push(dataAtual);
      Logger.log('Bloco unido! Agora vai de ' + formatarData(blocoAtual.inicio) + ' a ' + formatarData(blocoAtual.fim) +
        ' (' + contarDiasCalendario(blocoAtual.inicio, blocoAtual.fim) + ' dias de calendário)');
    } else {
      blocos.push(blocoAtual);
      Logger.log('Não consecutivo! A fechar bloco de ' + contarDiasCalendario(blocoAtual.inicio, blocoAtual.fim) + ' dia(s)');

      blocoAtual = {
        inicio: dataAtual,
        fim: dataAtual,
        dias: [dataAtual]
      };
    }
  }

  blocos.push(blocoAtual);

  return blocos;
}

/**
 * Formata a data para string legível (DD/MM/YYYY).
 */
function formatarData(data) {
  const dia = String(data.getDate()).padStart(2, '0');
  const mes = String(data.getMonth() + 1).padStart(2, '0');
  const ano = data.getFullYear();
  return dia + '/' + mes + '/' + ano;
}

/**
 * Obtém o calendário do utilizador.
 */
function obterCalendario() {
  try {
    const calendarioPrincipal = CalendarApp.getDefaultCalendar();
    const nomeConfigurado = String(CONFIG.CALENDARIO.NOME || '').trim();

    if (!nomeConfigurado) {
      return calendarioPrincipal;
    }

    if (calendarioPrincipal.getName() === nomeConfigurado) {
      return calendarioPrincipal;
    }

    const calendarios = CalendarApp.getAllOwnedCalendars();
    for (const cal of calendarios) {
      if (cal.getName() === nomeConfigurado) {
        return cal;
      }
    }

    Logger.log('Calendário "' + nomeConfigurado + '" não encontrado. A usar calendário principal.');
    return calendarioPrincipal;

  } catch (erro) {
    Logger.log('Erro ao obter calendário: ' + erro.message);
    return null;
  }
}

/**
 * Indica se um evento do Calendar foi gerado por este script: título "Férias"
 * (incluindo "Férias (X dias)") ou marcador presente na descrição.
 */
function eventoGeradoPeloScript(evento) {
  const titulo = evento.getTitle();
  const descricao = evento.getDescription() || '';
  return titulo.startsWith(CONFIG.CALENDARIO.TITULO_EVENTO) || descricao.indexOf(CONFIG.CALENDARIO.MARCADOR) !== -1;
}

/**
 * Devolve os eventos gerados pelo script no ano indicado, sem os apagar.
 * Uma única leitura ao Calendar por sincronização (o mesmo custo de antes);
 * a decisão do que criar, atualizar ou remover fica a cargo de sincronizarBlocosComDiferenca.
 */
function obterEventosGeradosNoAno(calendario, ano) {
  const dataInicio = new Date(ano, 0, 1);  // 1 de janeiro.
  const dataFim = new Date(ano + 1, 0, 1); // 1 de janeiro do ano seguinte; inclui 31 de dezembro.
  return calendario.getEvents(dataInicio, dataFim).filter(eventoGeradoPeloScript);
}

/**
 * Sincroniza os eventos gerados pelo script por diferença: cria só os blocos em falta,
 * atualiza só os que mudaram (título, descrição ou datas) e apaga só os que já não
 * correspondem a nenhum bloco pintado. Substitui a estratégia anterior de apagar todos
 * os eventos do ano e recriá-los do zero a cada sincronização, que gastava a quota diária
 * do Calendar sem necessidade sempre que nada tinha realmente mudado.
 * Cada bloco é isolado num try/catch: um erro num bloco (ex. um evento antigo inválido)
 * fica registado e contado em `falhados`, mas não aborta os restantes blocos da mesma
 * sincronização. Um erro de quota é a única exceção — é propagado ao chamador, que trata
 * o bloqueio e a retentativa (não faz sentido continuar a tentar quando a quota esgotou).
 * Aplica CONFIG.SYNC.MAX_EVENT_MUTATIONS_PER_RUN como limite defensivo por execução;
 * devolve `pendente: true` se o limite foi atingido, para o chamador agendar uma continuação.
 */
function sincronizarBlocosComDiferenca(calendario, eventosDesejados, eventosExistentes) {
  const porChave = new Map();
  eventosExistentes.forEach(evento => {
    const chave = String(evento.getStartTime().getTime());
    if (!porChave.has(chave)) {
      porChave.set(chave, evento);
    }
  });

  let criados = 0;
  let atualizados = 0;
  let removidos = 0;
  let falhados = 0;
  let mutacoes = 0;
  const usados = new Set();
  const limiteAtingido = () => mutacoes >= CONFIG.SYNC.MAX_EVENT_MUTATIONS_PER_RUN;

  function registarFalha(contexto, erro) {
    const mensagem = erro && erro.message ? erro.message : String(erro);
    if (REGEX_QUOTA_CALENDAR.test(mensagem)) {
      throw erro; // Quota esgotada: não continuar a tentar; o chamador trata o backoff.
    }
    falhados++;
    mutacoes++;
    Logger.log('Erro ao sincronizar bloco (' + contexto + '): ' + mensagem);
  }

  eventosDesejados.forEach(desejado => {
    if (limiteAtingido()) return;
    const existente = porChave.get(desejado.chave);

    try {
      if (!existente) {
        calendario.createAllDayEvent(desejado.titulo, desejado.inicio, desejado.fimApi, { description: desejado.descricao });
        criados++;
        mutacoes++;
        return;
      }

      usados.add(existente);

      // Eventos de dia inteiro não têm um "setter" de intervalo multi-dia no CalendarApp;
      // quando a duração do bloco muda, o evento é recriado, mas só esse bloco (não o ano inteiro).
      if (existente.getEndTime().getTime() !== desejado.fimApi.getTime()) {
        existente.deleteEvent();
        calendario.createAllDayEvent(desejado.titulo, desejado.inicio, desejado.fimApi, { description: desejado.descricao });
        atualizados++;
        mutacoes++;
        return;
      }

      let alterado = false;
      if (existente.getTitle() !== desejado.titulo) {
        existente.setTitle(desejado.titulo);
        alterado = true;
      }
      if (existente.getDescription() !== desejado.descricao) {
        existente.setDescription(desejado.descricao);
        alterado = true;
      }
      if (alterado) {
        atualizados++;
        mutacoes++;
      }
    } catch (erroBloco) {
      registarFalha(formatarData(desejado.inicio), erroBloco);
    }
  });

  eventosExistentes.forEach(evento => {
    if (limiteAtingido()) return;
    if (!usados.has(evento)) {
      try {
        evento.deleteEvent();
        removidos++;
        mutacoes++;
      } catch (erroRemocao) {
        registarFalha('remoção de evento antigo', erroRemocao);
      }
    }
  });

  return { criados, atualizados, removidos, falhados, pendente: limiteAtingido() };
}

// GESTÃO DE TRIGGERS (AUTOMAÇÃO)
// ============================

/**
 * Instala trigger para atualizar automaticamente quando a folha é editada.
 * Nota: este trigger não deteta mudanças de cores, apenas de valores.
 */
function instalarTrigger() {
  try {
    // Remover triggers existentes da mesma função, para evitar duplicação.
    removerTriggersExistentes();

    // Criar novo trigger onEdit, para mudanças de valores.
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
 * Instala triggers para sincronização automática:
 * - diariamente, como rede de segurança;
 * - quando há alterações de formatação/cor no ficheiro (contadores imediatos,
 *   Calendar agendado com debounce — ver onAlteracaoPlanilha/agendarSincronizacaoCalendar).
 * Também limpa qualquer bloqueio de quota do Calendar herdado de uma execução anterior.
 */
function instalarTriggerAutomatico() {
  try {
    // Remover triggers automáticos existentes.
    removerTriggersAutomaticos();
    PropertiesService.getScriptProperties().deleteProperty(PROPRIEDADES.quotaRetryAt);

    // Criar trigger diário, como rede de segurança (a frescura normal vem do debounce ao pintar).
    ScriptApp.newTrigger(HANDLERS.sincronizacaoDiaria)
      .timeBased()
      .everyDays(1)
      .atHour(3)
      .create();

    ScriptApp.newTrigger(HANDLERS.alteracaoFolha)
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onChange()
      .create();

    Logger.log('Triggers automáticos instalados (diário às 3h + alterações de cor/formato)');
    mostrarNotificacao(
      'Sincronização automática ativada! Atualiza ao pintar (com pequeno atraso) e diariamente.',
      'Automação Total Ativa',
      5
    );

  } catch (erro) {
    Logger.log('❌ Erro ao instalar trigger automático: ' + erro.message);
    mostrarNotificacao('Erro ao ativar sincronização automática.', 'Erro', 5);
  }
}

/**
 * Remove o trigger de atualização automática por edição.
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
 * Remove os triggers de sincronização automática.
 */
function removerTriggerAutomatico() {
  try {
    removerTriggersAutomaticos();

    Logger.log('✅ Triggers automáticos (diário, onChange e Calendar pendente) removidos');
    mostrarNotificacao('Sincronização automática desativada!', 'Automação Total Desativada', 3);

  } catch (erro) {
    Logger.log('❌ Erro ao remover trigger automático: ' + erro.message);
  }
}

/**
 * Remove todos os triggers existentes da função atualizarContadores.
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
 * Remove todos os triggers automáticos (sincronização diária, onChange e Calendar pendente/debounce).
 */
function removerTriggersAutomaticos() {
  const handlers = {
    sincronizarTudo: true,
    onAlteracaoPlanilha: true,
    sincronizarCalendarPendente: true
  };
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (handlers[trigger.getHandlerFunction()]) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// ============================
// MENU PERSONALIZADO
// ============================

/**
 * Cria menu personalizado ao abrir o Google Sheet.
 * Executado automaticamente pelo trigger onOpen.
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
    .addItem('🤖 Ativar Sincronização Automática', 'instalarTriggerAutomatico')
    .addSeparator()
    .addItem('❌ Desativar Atualização ao Editar', 'removerTrigger')
    .addItem('⛔ Desativar Sincronização Automática', 'removerTriggerAutomatico')
    .addSeparator()
    .addItem('🔍 Testar Deteção de Cores', 'testarDetecaoCores')
    .addItem('ℹ️ Ajuda', 'mostrarAjuda')
    .addToUi();
}

/**
 * Mostra janela de ajuda com instruções.
 */
function mostrarAjuda() {
  const ui = SpreadsheetApp.getUi();

  const mensagem = [
    'GESTÃO DE FÉRIAS ' + obterAnoDaSheet(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()),
    '',
    'SINCRONIZAR TUDO (RECOMENDADO)',
    '- Menu: "SINCRONIZAR TUDO"',
    '- Atualiza contadores e sincroniza Calendar',
    '- Usa sempre que pintares células de férias',
    '',
    'SINCRONIZAÇÃO AUTOMÁTICA',
    '- Menu: "Ativar Sincronização Automática"',
    '- Contadores atualizam ao pintar; o Calendar sincroniza pouco depois (agrega várias pinturas seguidas)',
    '- Trigger diário como rede de segurança adicional',
    '- Pinta as células à vontade e o sistema trata da sincronização',
    '',
    'CONTADORES MANUAIS',
    '- Menu: "Atualizar Contadores": só números',
    '- Menu: "Sincronizar com Calendar": só eventos',
    '',
    'CORES A USAR',
    '- Roxo (#d9d2e9): férias planeadas',
    '- Verde (#d9ead3): dia de aniversário da empresa',
    '',
    'QUOTA DO GOOGLE CALENDAR',
    '- Se a quota diária do Calendar esgotar, o sistema pausa sozinho e tenta novamente mais tarde',
    '- Para forçar uma tentativa imediata, usa "Sincronizar com Calendar" ou "SINCRONIZAR TUDO"',
    '',
    'RECOMENDAÇÕES',
    'OPÇÃO 1: totalmente automático:',
    '1. Ativa "Sincronização Automática"',
    '2. Pinta células à vontade',
    '3. Aguarda alguns minutos',
    '4. Tudo atualiza sozinho',
    '',
    'OPÇÃO 2: semiautomático:',
    '1. Pinta as células de férias',
    '2. Clica em "SINCRONIZAR TUDO"',
    '3. Pronto! (instantâneo)',
    '',
    'Se houver problemas, usa "Testar Deteção de Cores"',
  ].join('\n');

  ui.alert('Ajuda - Gestão de Férias', mensagem, ui.ButtonSet.OK);
}

// ============================
// CONFIGURAÇÃO INICIAL
// ============================

/**
 * Configura o Google Sheet pela primeira vez.
 * Cria a estrutura de legenda e contadores.
 * Executar uma vez, após colar o código.
 */
function configurarSheet() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Estrutura da legenda, a partir da linha 18, coluna B.
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

    comSupressaoAlteracaoFolha(function () {
      // Inserir dados na linha 18, coluna B.
      sheet.getRange(18, 2, legenda.length, 2).setValues(legenda);

      // Aplicar formatação às células da legenda.
      sheet.getRange('B18:B23').setFontWeight('normal');
      sheet.getRange('B18').setFontWeight('bold'); // Destaque da primeira linha.
      sheet.getRange('B24').setFontWeight('bold'); // Separador antes do aniversário.

      // Aplicar cores às células da legenda, como exemplos visuais.
      sheet.getRange('B18:C18').setBackground('#e6b8af'); // Disponíveis do ano corrente.
      sheet.getRange('B19:C19').setBackground('#fff2cc'); // Dias transitados do ano anterior.
      sheet.getRange('B20:C20').setBackground('#d9d2e9'); // Gozadas.
      sheet.getRange('B21:C21').setBackground('#d9d2e9'); // Planeadas.
      sheet.getRange('B22:C22').setBackground('#d9d2e9'); // Total.
      sheet.getRange('B23:C23').setBackground('#b6d7a8'); // Restantes.
      sheet.getRange('B25:C27').setBackground('#d9ead3'); // Aniversário.
    });

    Logger.log('✅ Sheet configurado com sucesso!');
    mostrarNotificacao('Legenda e contadores configurados!', 'Configuração completa', 3);

    // Executar atualização inicial dos contadores.
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
 * Mostra notificação toast no Google Sheets.
 */
function mostrarNotificacao(mensagem, titulo, duracao) {
  SpreadsheetApp.getActiveSpreadsheet().toast(mensagem, titulo, duracao);
}

/**
 * Função de teste para validar cores.
 * Executa e verifica o log para ver as cores detetadas.
 * 
 * COMO USAR:
 * 1. No Apps Script Editor, seleciona esta função no dropdown.
 * 2. Clica em Executar (▶️).
 * 3. Vai a "Execuções" (ou View > Logs).
 * 4. Vê todas as cores encontradas no teu calendário.
 */
function testarDetecaoCores() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(CONFIG.CALENDAR_RANGE);
  const valores = range.getValues();
  const cores = range.getBackgrounds();

  const coresEncontradas = new Map(); // Guarda cor e quantidade de ocorrências.

  // Percorrer todas as células.
  for (let i = 0; i < cores.length; i++) {
    for (let j = 0; j < cores[i].length; j++) {
      const cor = cores[i][j];
      const valor = valores[i][j];

      // Ignorar branco, preto, células vazias ou com texto.
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

  // Mostrar resultados no log.
  Logger.log('═══════════════════════════════════════');
  Logger.log('🎨 DIAGNÓSTICO DE CORES - CALENDÁRIO ' + obterAnoDaSheet(sheet));
  Logger.log('═══════════════════════════════════════\n');

  Logger.log('📊 Cores encontradas no calendário (apenas células com números):');
  coresEncontradas.forEach((quantidade, cor) => {
    Logger.log('  ' + cor + ' : ' + quantidade + ' células');
  });

  Logger.log('\n🎯 Cores configuradas para deteção:');
  Logger.log('  Férias (ano corrente): ' + CONFIG.CORES.FERIAS_ATUAL.toLowerCase());
  Logger.log('  Férias (ano corrente - alt): ' + CONFIG.CORES.FERIAS_ATUAL_ALT.toLowerCase());
  Logger.log('  Férias (ano anterior): ' + CONFIG.CORES.FERIAS_ANTERIOR.toLowerCase());
  Logger.log('  Aniversário (verde): ' + CONFIG.CORES.ANIVERSARIO.toLowerCase());

  Logger.log('\n✅ Correspondências encontradas:');
  let encontrouFerias = false;
  let encontrouAniversario = false;

  coresEncontradas.forEach((quantidade, cor) => {
    if (cor === CONFIG.CORES.FERIAS_ATUAL.toLowerCase()) {
      Logger.log('Férias (ano corrente): ' + quantidade + ' células detetadas!');
      encontrouFerias = true;
    }
    if (cor === CONFIG.CORES.FERIAS_ATUAL_ALT.toLowerCase()) {
      Logger.log('Férias (ano corrente - alt): ' + quantidade + ' células detetadas!');
      encontrouFerias = true;
    }
    if (cor === CONFIG.CORES.FERIAS_ANTERIOR.toLowerCase()) {
      Logger.log('Férias (ano anterior): ' + quantidade + ' células detetadas!');
      encontrouFerias = true;
    }
    if (cor === CONFIG.CORES.ANIVERSARIO.toLowerCase()) {
      Logger.log('Aniversário: ' + quantidade + ' células verdes detetadas!');
      encontrouAniversario = true;
    }
  });

  if (!encontrouFerias) {
    Logger.log('  ✗ Nenhuma célula roxa de férias encontrada!');
    Logger.log('    → Verifica se usaste a cor: ' + CONFIG.CORES.FERIAS_ATUAL);
  }

  if (!encontrouAniversario) {
    Logger.log('  ✗ Nenhuma célula verde de aniversário encontrada!');
    Logger.log('    → Verifica se usaste a cor: ' + CONFIG.CORES.ANIVERSARIO);
  }

  Logger.log('\n═══════════════════════════════════════');

  // Mostrar também como notificação.
  let mensagem = 'Cores encontradas: ' + coresEncontradas.size + '\n';
  mensagem += encontrouFerias ? '✓ Férias OK\n' : '✗ Férias não encontradas\n';
  mensagem += encontrouAniversario ? '✓ Aniversário OK' : '✗ Aniversário não encontrado';

  mostrarNotificacao(mensagem, 'Teste de Cores', 8);
}

/**
 * Atualiza as cores no CONFIG, com base no que está pintado na folha.
 * Executar depois de testarDetecaoCores(), se as cores não corresponderem.
 */
function atualizarCoresAutomaticamente() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(CONFIG.CALENDAR_RANGE);
  const valores = range.getValues();
  const cores = range.getBackgrounds();

  const coresEncontradas = new Map();

  // Contar ocorrências de cada cor.
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

  // Ordenar por quantidade, com as mais usadas primeiro.
  const coresOrdenadas = Array.from(coresEncontradas.entries()).sort((a, b) => b[1] - a[1]);

  if (coresOrdenadas.length >= 1) {
    const corMaisUsada = coresOrdenadas[0][0];
    Logger.log('Sugestão: definir cor de férias como: ' + corMaisUsada);
    Logger.log('  (Encontradas ' + coresOrdenadas[0][1] + ' células com esta cor)');
  }

  if (coresOrdenadas.length >= 2) {
    const segundaCorMaisUsada = coresOrdenadas[1][0];
    Logger.log('Sugestão: definir cor de aniversário como: ' + segundaCorMaisUsada);
    Logger.log('  (Encontradas ' + coresOrdenadas[1][1] + ' células com esta cor)');
  }

  Logger.log('\n💡 Para atualizar as cores no código:');
  Logger.log('1. Edita o objeto CONFIG no topo do código');
  Logger.log('2. Altera os valores em CONFIG.CORES.FERIAS_ATUAL, CONFIG.CORES.FERIAS_ATUAL_ALT, CONFIG.CORES.FERIAS_ANTERIOR ou CONFIG.CORES.ANIVERSARIO');
  Logger.log('3. Guarda o script e executa atualizarContadores() novamente');
}
