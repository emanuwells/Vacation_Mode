const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const ANO = 2026;
const COR_FERIAS = '#d9d2e9'; // Tem de corresponder a CONFIG.CORES.FERIAS_ATUAL em Vacation_Mode.js.

function construirGrelha(pintadas) {
  const valores = [];
  const cores = [];
  for (let mes = 0; mes < 12; mes++) {
    const linhaValores = [];
    const linhaCores = [];
    for (let dia = 1; dia <= 31; dia++) {
      linhaValores.push(dia);
      linhaCores.push(pintadas.has(mes + '-' + dia) ? COR_FERIAS : '#ffffff');
    }
    valores.push(linhaValores);
    cores.push(linhaCores);
  }
  return { valores, cores };
}

function criarSheet(nome, pintadas, feriasRestantes) {
  const celulas = new Map([['C23', feriasRestantes]]);
  return {
    getName: () => nome,
    getRange(a1) {
      if (a1 === 'C5:AM16') {
        const { valores, cores } = construirGrelha(pintadas);
        return { getValues: () => valores, getBackgrounds: () => cores, getRow: () => 5, getColumn: () => 3 };
      }
      return { getValue: () => celulas.get(a1) || 0, setValue: v => celulas.set(a1, v) };
    }
  };
}

function criarContexto() {
  const eventos = [];
  const chamadas = { getEvents: 0, createAllDayEvent: 0, deleteEvent: 0, setTitle: 0, setDescription: 0 };
  const triggersCriados = [];
  const triggersEliminados = [];
  const propriedadesScript = new Map();
  const propriedadesDoc = new Map();
  let triggersExistentes = [];
  let lancarErroQuota = false;

  function criarEvento(titulo, inicio, fim, descricao) {
    const evento = {
      titulo,
      inicio: new Date(inicio),
      fim: new Date(fim),
      descricao,
      apagado: false,
      getTitle: () => evento.titulo,
      getDescription: () => evento.descricao,
      getStartTime: () => evento.inicio,
      getEndTime: () => evento.fim,
      setTitle(t) { chamadas.setTitle++; evento.titulo = t; return evento; },
      setDescription(d) { chamadas.setDescription++; evento.descricao = d; return evento; },
      deleteEvent() { chamadas.deleteEvent++; evento.apagado = true; }
    };
    eventos.push(evento);
    return evento;
  }

  const calendario = {
    getName: () => 'Calendário de Teste',
    getEvents(inicio, fim) {
      chamadas.getEvents++;
      if (lancarErroQuota) {
        throw new Error('Service invoked too many times for one day: calendar.');
      }
      return eventos.filter(ev => !ev.apagado && ev.inicio < fim && ev.fim > inicio);
    },
    createAllDayEvent(titulo, inicio, fim, opts) {
      chamadas.createAllDayEvent++;
      return criarEvento(titulo, inicio, fim, opts.description);
    }
  };

  function novoTrigger(handler) {
    const criado = { handler };
    const builder = {
      timeBased() { criado.tipo = 'timeBased'; return builder; },
      everyDays(d) { criado.everyDays = d; return builder; },
      atHour(h) { criado.atHour = h; return builder; },
      after(ms) { criado.after = ms; return builder; },
      forSpreadsheet(ss) { criado.spreadsheet = ss; return builder; },
      onChange() { criado.tipo = 'onChange'; return builder; },
      onEdit() { criado.tipo = 'onEdit'; return builder; },
      create() { triggersCriados.push(criado); return criado; }
    };
    return builder;
  }

  const scriptProperties = {
    getProperty: chave => propriedadesScript.get(chave) ?? null,
    setProperty(chave, valor) { propriedadesScript.set(chave, valor); return scriptProperties; },
    deleteProperty(chave) { propriedadesScript.delete(chave); return scriptProperties; }
  };
  const documentProperties = {
    getProperty: chave => propriedadesDoc.get(chave) ?? null,
    setProperty(chave, valor) { propriedadesDoc.set(chave, valor); return documentProperties; },
    deleteProperty(chave) { propriedadesDoc.delete(chave); return documentProperties; }
  };

  const spreadsheet = {
    getUrl: () => 'https://docs.google.com/spreadsheets/d/teste',
    toast() {},
    getSheets: () => spreadsheet._sheets || [],
    getActiveSheet: () => (spreadsheet._sheets || [])[0]
  };

  const contexto = {
    CalendarApp: {
      getDefaultCalendar: () => calendario,
      getAllOwnedCalendars: () => [calendario]
    },
    LockService: {
      getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }),
      getScriptLock: () => ({ tryLock: () => true, releaseLock() {} })
    },
    Logger: { log() {} },
    PropertiesService: {
      getScriptProperties: () => scriptProperties,
      getDocumentProperties: () => documentProperties
    },
    ScriptApp: {
      getProjectTriggers: () => triggersExistentes,
      deleteTrigger: item => triggersEliminados.push(item),
      newTrigger: novoTrigger
    },
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    Utilities: { sleep() {} }
  };

  vm.createContext(contexto);
  const script = fs.readFileSync(path.join(__dirname, '..', 'src', 'Vacation_Mode.js'), 'utf8');
  vm.runInContext(script, contexto);

  return {
    contexto,
    calendario,
    chamadas,
    eventos,
    eventosAtivos: () => eventos.filter(e => !e.apagado),
    propriedadesScript,
    triggersCriados,
    setSheets: sheets => { spreadsheet._sheets = sheets; },
    setLancarErroQuota: valor => { lancarErroQuota = valor; },
    setTriggersExistentes: triggers => { triggersExistentes = triggers; }
  };
}

// --- Cenário 1: sincronizar sem alterações não deve tocar no Calendar. ---
const c1 = criarContexto();
const pintadasA = new Set(['7-12', '7-13']); // Agosto, 12-13 (quarta-quinta, sem extensão de fim de semana).
const sheetA = criarSheet('Calendário 2026', pintadasA, 7);
c1.setSheets([sheetA]);

c1.contexto.sincronizarComCalendar(sheetA, ANO, true);
assert.equal(c1.chamadas.createAllDayEvent, 1, 'primeira sincronização deve criar 1 evento');
assert.equal(c1.eventosAtivos().length, 1);

c1.contexto.sincronizarComCalendar(sheetA, ANO, true);
assert.equal(c1.chamadas.createAllDayEvent, 1, 'sem alterações, não deve criar de novo');
assert.equal(c1.chamadas.deleteEvent, 0, 'sem alterações, não deve apagar nada');
assert.equal(c1.chamadas.setTitle, 0, 'sem alterações, não deve atualizar o título');
assert.equal(c1.chamadas.setDescription, 0, 'sem alterações, não deve atualizar a descrição');
console.log('OK: sincronizar sem alterações não cria nem apaga nenhum evento (corrige o bug de apagar tudo/recriar tudo).');

// --- Cenário 2: dois blocos; crescer um bloco só afeta esse bloco. ---
const c2 = criarContexto();
const pintadasB = new Set(['7-12', '7-13', '8-2', '8-3']); // Bloco A: Ago 12-13. Bloco B: Set 2-3.
const sheetB = criarSheet('Calendário 2026', pintadasB, 5);
c2.setSheets([sheetB]);

c2.contexto.sincronizarComCalendar(sheetB, ANO, true);
assert.equal(c2.chamadas.createAllDayEvent, 2, 'deve criar os 2 blocos pintados');
const eventoBlocoB = c2.eventosAtivos().find(e => e.inicio.getMonth() === 8);
assert.ok(eventoBlocoB, 'o bloco de setembro deve existir');

// Pintar mais um dia a seguir ao bloco A (Ago 14, sexta) — o bloco cresce, não recomeça do zero.
pintadasB.add('7-14');
c2.contexto.sincronizarComCalendar(sheetB, ANO, true);
assert.equal(c2.chamadas.createAllDayEvent, 3, 'deve criar apenas 1 evento novo para o bloco que cresceu');
assert.equal(c2.chamadas.deleteEvent, 1, 'deve apagar apenas o evento antigo do bloco que cresceu');
assert.equal(c2.chamadas.setTitle, 0);
assert.equal(c2.chamadas.setDescription, 0);
assert.equal(c2.eventosAtivos().length, 2, 'continuam a existir exatamente 2 eventos (bloco A crescido + bloco B intacto)');
assert.equal(eventoBlocoB.apagado, false, 'o bloco de setembro não deve ter sido tocado');
console.log('OK: alargar um bloco só recria esse bloco; os restantes ficam intactos.');

// --- Cenário 3: despintar um bloco só apaga o evento desse bloco. ---
pintadasB.delete('8-2');
pintadasB.delete('8-3');
c2.contexto.sincronizarComCalendar(sheetB, ANO, true);
assert.equal(c2.chamadas.deleteEvent, 2, 'deve apagar apenas o evento do bloco despintado');
assert.equal(c2.chamadas.createAllDayEvent, 3, 'não deve criar nada ao despintar');
assert.equal(c2.eventosAtivos().length, 1, 'só o bloco de agosto continua ativo');
console.log('OK: despintar um bloco apaga só esse evento, sem tocar nos restantes.');

// --- Cenário 4: quota esgotada numa folha não deve ser tentada novamente noutra folha na mesma execução. ---
const c4 = criarContexto();
const sheet2026 = criarSheet('Calendário 2026', new Set(['7-12', '7-13']), 7);
const sheet2025 = criarSheet('Calendário 2025', new Set(['7-12', '7-13']), 7);
c4.setSheets([sheet2026, sheet2025]);
c4.setLancarErroQuota(true);

c4.contexto.sincronizarTudo({ automatico: true });

assert.equal(c4.chamadas.getEvents, 1, 'só a primeira folha deve chegar a tentar o Calendar nesta execução');
assert.equal(c4.chamadas.createAllDayEvent, 0);
assert.ok(Number(c4.propriedadesScript.get('VACATION_MODE_CALENDAR_QUOTA_RETRY_AT')) > Date.now(),
  'deve gravar uma retentativa futura para a quota diária');
assert.ok(c4.triggersCriados.some(t => t.handler === 'sincronizarCalendarPendente'),
  'deve agendar uma retentativa automática');
console.log('OK: quota esgotada numa folha bloqueia as restantes na mesma execução e agenda retentativa.');

// Execução manual limpa o bloqueio e força uma tentativa real.
c4.setLancarErroQuota(false);
c4.contexto.sincronizarComCalendar(sheet2026, ANO, false);
assert.equal(c4.propriedadesScript.has('VACATION_MODE_CALENDAR_QUOTA_RETRY_AT'), false,
  'execução manual deve limpar o bloqueio de quota');
assert.equal(c4.chamadas.getEvents, 2, 'a execução manual deve mesmo tentar o Calendar');
assert.equal(c4.chamadas.createAllDayEvent, 1, 'com a quota livre, o evento é finalmente criado');
console.log('OK: uma sincronização manual limpa o bloqueio de quota e força uma tentativa real.');

// --- Cenário 5: um erro num bloco não deve abortar os restantes (isolamento por bloco). ---
const c5 = criarContexto();
const pintadasC = new Set(['7-12', '7-13', '8-2', '8-3']); // Bloco A: Ago 12-13. Bloco B: Set 2-3.
const sheetC = criarSheet('Calendário 2026', pintadasC, 4);
c5.setSheets([sheetC]);

const datasC = c5.contexto.obterDatasFerias(sheetC, ANO);
const blocosC = c5.contexto.agruparDatasConsecutivas(datasC);
const desejadosC = blocosC.map(b => c5.contexto.construirEventoDesejado(b, 4, 'https://teste'));
assert.equal(desejadosC.length, 2, 'preparação do cenário: devem existir 2 blocos desejados');

let tentativas = 0;
const calendarioComFalhaPontual = {
  createAllDayEvent(titulo, inicio, fim, opts) {
    tentativas++;
    if (tentativas === 1) {
      throw new Error('Falha simulada só no primeiro bloco (ex.: evento antigo inválido).');
    }
    return c5.calendario.createAllDayEvent(titulo, inicio, fim, opts);
  }
};

const resultadoC = c5.contexto.sincronizarBlocosComDiferenca(calendarioComFalhaPontual, desejadosC, []);
assert.equal(resultadoC.criados, 1, 'o segundo bloco deve ser criado apesar do erro no primeiro');
assert.equal(resultadoC.falhados, 1, 'o bloco com erro deve ficar registado como falhado, não abortar tudo');
assert.equal(tentativas, 2, 'ambos os blocos devem ser tentados');
console.log('OK: um erro pontual num bloco não aborta os restantes (só esse bloco falha; os outros sincronizam à mesma).');

// Um erro de quota, em contraste, interrompe de imediato — não faz sentido continuar a gastar quota já esgotada.
tentativas = 0;
const calendarioComQuotaEsgotada = {
  createAllDayEvent() {
    tentativas++;
    throw new Error('Service invoked too many times for one day: calendar.');
  }
};
assert.throws(
  () => c5.contexto.sincronizarBlocosComDiferenca(calendarioComQuotaEsgotada, desejadosC, []),
  /Service invoked too many times/
);
assert.equal(tentativas, 1, 'um erro de quota não deve ser tentado novamente para os restantes blocos desta chamada');
console.log('OK: um erro de quota interrompe de imediato a sincronização por diferença, sem tentar os restantes blocos.');
