const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const criados = [];
const eliminados = [];
const propriedadesScript = new Map();
const propriedadesDoc = new Map();
const folha = { toast() {} };
let triggersExistentes = [];
let locksScriptAtivos = 0;
let calendarCalls = 0;

function trigger(handler) {
  return { getHandlerFunction: () => handler };
}

function novoTrigger(handler) {
  const criado = { handler };
  const builder = {
    timeBased() {
      criado.tipo = 'timeBased';
      return builder;
    },
    everyDays(dias) {
      criado.everyDays = dias;
      return builder;
    },
    atHour(hora) {
      criado.atHour = hora;
      return builder;
    },
    after(atrasoMs) {
      criado.after = atrasoMs;
      return builder;
    },
    forSpreadsheet(spreadsheet) {
      assert.equal(spreadsheet, folha);
      criado.spreadsheet = spreadsheet;
      return builder;
    },
    onChange() {
      criado.tipo = 'onChange';
      return builder;
    },
    onEdit() {
      criado.tipo = 'onEdit';
      return builder;
    },
    create() {
      criados.push(criado);
      return criado;
    }
  };
  return builder;
}

const scriptProperties = {
  getProperty: chave => propriedadesScript.get(chave) ?? null,
  setProperty(chave, valor) {
    propriedadesScript.set(chave, valor);
    return scriptProperties;
  },
  deleteProperty(chave) {
    propriedadesScript.delete(chave);
    return scriptProperties;
  }
};

const documentProperties = {
  getProperty: chave => propriedadesDoc.get(chave) ?? null,
  setProperty(chave, valor) {
    propriedadesDoc.set(chave, valor);
    return documentProperties;
  },
  deleteProperty(chave) {
    propriedadesDoc.delete(chave);
    return documentProperties;
  }
};

const contexto = {
  CalendarApp: {
    getDefaultCalendar() {
      calendarCalls++;
      throw new Error('CalendarApp não deve ser chamado ao agendar/debouncer uma edição.');
    },
    getAllOwnedCalendars() {
      calendarCalls++;
      throw new Error('CalendarApp não deve ser chamado ao agendar/debouncer uma edição.');
    }
  },
  LockService: {
    getScriptLock: () => ({
      tryLock() {
        locksScriptAtivos++;
        return true;
      },
      releaseLock() {
        locksScriptAtivos--;
      }
    }),
    getDocumentLock: () => ({
      tryLock: () => true,
      releaseLock() {}
    })
  },
  Logger: { log() {} },
  PropertiesService: {
    getScriptProperties: () => scriptProperties,
    getDocumentProperties: () => documentProperties
  },
  ScriptApp: {
    getProjectTriggers: () => triggersExistentes,
    deleteTrigger: item => eliminados.push(item),
    newTrigger: novoTrigger
  },
  SpreadsheetApp: {
    getActiveSpreadsheet: () => folha,
    getUi: () => ({ createMenu: () => ({ addItem: () => ({}) }) })
  },
  Utilities: { sleep() {} }
};

vm.createContext(contexto);
const script = fs.readFileSync(path.join(__dirname, '..', 'src', 'Vacation_Mode.js'), 'utf8');
vm.runInContext(script, contexto);

// --- instalarTriggerAutomatico: trigger diário (não 5 em 5 minutos) + onChange, e limpa bloqueio de quota. ---
const triggerDiarioAntigo = trigger('sincronizarTudo');
const triggerAlteracaoAntigo = trigger('onAlteracaoPlanilha');
const triggerPendenteAntigo = trigger('sincronizarCalendarPendente');
const triggerAlheio = trigger('outraFuncao');
triggersExistentes = [triggerDiarioAntigo, triggerAlteracaoAntigo, triggerPendenteAntigo, triggerAlheio];
propriedadesScript.set('VACATION_MODE_CALENDAR_QUOTA_RETRY_AT', String(Date.now() + 60 * 60 * 1000));

contexto.instalarTriggerAutomatico();

assert.deepEqual(eliminados, [triggerDiarioAntigo, triggerAlteracaoAntigo, triggerPendenteAntigo]);
assert.deepEqual(criados, [
  { handler: 'sincronizarTudo', tipo: 'timeBased', everyDays: 1, atHour: 3 },
  { handler: 'onAlteracaoPlanilha', tipo: 'onChange', spreadsheet: folha }
]);
assert.equal(propriedadesScript.has('VACATION_MODE_CALENDAR_QUOTA_RETRY_AT'), false);
console.log('OK: instalarTriggerAutomatico cria trigger diário (não 5 em 5 min) + onChange e limpa a quota.');

// --- removerTriggerAutomatico: remove os três handlers automáticos, incluindo o novo "pendente". ---
criados.length = 0;
eliminados.length = 0;
triggersExistentes = [triggerDiarioAntigo, triggerAlteracaoAntigo, triggerPendenteAntigo, triggerAlheio];
contexto.removerTriggerAutomatico();
assert.deepEqual(eliminados, [triggerDiarioAntigo, triggerAlteracaoAntigo, triggerPendenteAntigo]);
console.log('OK: removerTriggerAutomatico remove sincronizarTudo, onAlteracaoPlanilha e sincronizarCalendarPendente.');

// --- onAlteracaoPlanilha: contadores atualizam sem tocar no Calendar; Calendar é debounced (1 trigger, não N). ---
criados.length = 0;
eliminados.length = 0;
triggersExistentes = [];
propriedadesDoc.clear();
folha.getSheets = () => [{
  getName: () => 'Calendário 2026',
  getRange(a1) {
    if (a1 === 'C5:AM16') {
      const linha = new Array(31).fill('');
      return { getValues: () => Array.from({ length: 12 }, () => linha.slice()), getBackgrounds: () => Array.from({ length: 12 }, () => new Array(31).fill('#ffffff')) };
    }
    return { getValue: () => 0, setValue() {} };
  }
}];

contexto.onAlteracaoPlanilha({ changeType: 'FORMAT' });
assert.equal(calendarCalls, 0, 'onAlteracaoPlanilha não deve chamar o CalendarApp diretamente.');
assert.deepEqual(criados, [{ handler: 'sincronizarCalendarPendente', tipo: 'timeBased', after: 5 * 60 * 1000 }]);

// Pintar outra vez, pouco depois: só deve haver 1 trigger "pendente" agendado (debounce), não 2.
criados.length = 0;
eliminados.length = 0;
triggersExistentes = [{ handler: 'sincronizarCalendarPendente', getHandlerFunction: () => 'sincronizarCalendarPendente' }];
contexto.onAlteracaoPlanilha({ changeType: 'FORMAT' });
assert.equal(calendarCalls, 0);
assert.equal(eliminados.length, 1, 'O trigger pendente anterior deve ser removido antes de agendar o novo (debounce).');
assert.deepEqual(criados, [{ handler: 'sincronizarCalendarPendente', tipo: 'timeBased', after: 5 * 60 * 1000 }]);
console.log('OK: pinturas sucessivas agregam-se num único trigger "pendente" do Calendar (debounce real).');

// --- Com a quota bloqueada, o agendamento nem sequer cria um novo trigger. ---
criados.length = 0;
eliminados.length = 0;
triggersExistentes = [];
propriedadesScript.set('VACATION_MODE_CALENDAR_QUOTA_RETRY_AT', String(Date.now() + 60 * 60 * 1000));
contexto.onAlteracaoPlanilha({ changeType: 'FORMAT' });
assert.deepEqual(criados, []);
assert.equal(calendarCalls, 0);
propriedadesScript.delete('VACATION_MODE_CALENDAR_QUOTA_RETRY_AT');
console.log('OK: com a quota bloqueada, a sincronização do Calendar não é reagendada nem chama o CalendarApp.');

// --- Tipos de alteração irrelevantes continuam a ser ignorados. ---
criados.length = 0;
contexto.onAlteracaoPlanilha({ changeType: 'INSERT_GRID' });
assert.deepEqual(criados, []);
console.log('OK: tipos de alteração irrelevantes continuam a ser ignorados.');
