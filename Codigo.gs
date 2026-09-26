var SHEET_ID = "11zY5ug39GTip2bdpFI7oPJwbn32QdUb7CX_ci51GEck";
var SECRET = "";

var UNITS = ['cm1','cm2','cg','orto_sjq','orto_sjs','onco'];
var UNIT_LABELS = {
  cm1: 'CM1',
  cm2: 'CM2',
  cg: 'CG',
  orto_sjq: 'Orto Sao Joaquim',
  orto_sjs: 'Orto Sao Jose',
  onco: 'Oncologia'
};

function doPost(e) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Historico') || ss.insertSheet('Historico');

  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ ok: false, error: 'payload invalido' });
  }

  if (SECRET && data.secret !== SECRET) {
    return jsonOut({ ok: false, error: 'unauthorized' });
  }

  ensureHeaders(sheet);

  var u = data.units || {};
  function f(key, field) { return (u[key] && u[key][field] !== undefined) ? u[key][field] : ''; }
  function list(key, field) { return (u[key] && u[key][field]) ? u[key][field].join(' | ') : ''; }
  function count(key, field) { return (u[key] && u[key][field]) ? u[key][field].length : 0; }

  var row = [new Date(), data.data || '', data.hora || '', data.turno || '', data.nome || '', data.crm || '', data.cargo || ''];

  UNITS.forEach(function (k) {
    row.push(
      f(k, 'pacientes'), f(k, 'vagosM'), f(k, 'vagosF'), f(k, 'admissoes'), f(k, 'altas'),
      count(k, 'risco'), list(k, 'risco'),
      count(k, 'paliativos'), list(k, 'paliativos'),
      count(k, 'hemodialise'), list(k, 'hemodialise'),
      count(k, 'pendencias'), list(k, 'pendencias'),
      list(k, 'transferencias')
    );
  });

  var t = data.totais || {};
  row.push(
    t.pacientes || 0, t.vagosM || 0, t.vagosF || 0, t.vagos || 0,
    t.admissoes || 0, t.altas || 0, t.risco || 0, t.paliativos || 0, t.hemodialise || 0,
    t.pendencias || 0, t.ocupacao || 0
  );

  sheet.appendRow(row);

  return jsonOut({ ok: true });
}

function doGet(e) {
  return jsonOut({ ok: true, msg: 'Backend do Huddle ativo.' });
}

function ensureHeaders(sheet) {
  if (sheet.getLastRow() > 0) return;
  var headers = ['Carimbo de envio', 'Data', 'Hora', 'Turno', 'Preenchido por', 'CRM-CE', 'Funcao'];
  UNITS.forEach(function (k) {
    var L = UNIT_LABELS[k];
    headers.push(
      L + ' Pacientes', L + ' Vagos Masc', L + ' Vagos Fem', L + ' Admissoes', L + ' Altas',
      L + ' Risco (n)', L + ' Risco (lista)',
      L + ' Paliativos (n)', L + ' Paliativos (leitos)',
      L + ' Hemodialise (n)', L + ' Hemodialise (leitos)',
      L + ' Pendencias (n)', L + ' Pendencias (lista)',
      L + ' Transferencias'
    );
  });
  headers.push(
    'Total Pacientes', 'Total Vagos Masc', 'Total Vagos Fem', 'Total Vagos',
    'Total Admissoes', 'Total Altas', 'Total Risco', 'Total Paliativos', 'Total Hemodialise',
    'Total Pendencias', 'Ocupacao (%)'
  );
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#B3122A').setFontColor('#FFFFFF');
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function debugPost() {
  var fakeEvent = {
    postData: {
      contents: JSON.stringify({
        secret: '',
        data: '2026-09-26',
        hora: '23:10',
        turno: 'Noite',
        nome: 'TESTE DEBUG',
        crm: '0000',
        cargo: 'Medico Hospitalista',
        units: {
          cm1: { pacientes: '1', vagosM: '0', vagosF: '0', admissoes: '0', altas: '0', risco: [], paliativos: [], hemodialise: [], pendencias: [], transferencias: [] },
          cm2: { pacientes: '', vagosM: '', vagosF: '', admissoes: '', altas: '', risco: [], paliativos: [], hemodialise: [], pendencias: [], transferencias: [] },
          cg: { pacientes: '', vagosM: '', vagosF: '', admissoes: '', altas: '', risco: [], paliativos: [], hemodialise: [], pendencias: [], transferencias: [] },
          orto_sjq: { pacientes: '', vagosM: '', vagosF: '', admissoes: '', altas: '', risco: [], paliativos: [], hemodialise: [], pendencias: [], transferencias: [] },
          orto_sjs: { pacientes: '', vagosM: '', vagosF: '', admissoes: '', altas: '', risco: [], paliativos: [], hemodialise: [], pendencias: [], transferencias: [] },
          onco: { pacientes: '', vagosM: '', vagosF: '', admissoes: '', altas: '', risco: [], paliativos: [], hemodialise: [], pendencias: [], transferencias: [] }
        },
        totais: { pacientes: 1, vagosM: 0, vagosF: 0, vagos: 0, admissoes: 0, altas: 0, risco: 0, paliativos: 0, hemodialise: 0, pendencias: 0, ocupacao: 100 }
      })
    }
  };

  try {
    var result = doPost(fakeEvent);
    Logger.log('SUCESSO: ' + result.getContent());
  } catch (err) {
    Logger.log('ERRO CAPTURADO: ' + err.message);
    Logger.log('STACK: ' + err.stack);
  }
}
