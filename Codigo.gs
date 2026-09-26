/**
 * BACKEND DO HUDDLE — Painel de Leitos (Santa Casa de Misericórdia de Sobral)
 * ---------------------------------------------------------------------------
 * O que este script faz:
 *  - Recebe os dados enviados pelo app (uma requisição por envio)
 *  - Grava UMA linha nova por envio na aba "Histórico" (nunca sobrescreve)
 *  - Cria os cabeçalhos automaticamente na primeira vez
 *
 * COMO INSTALAR (uma única vez, funciona pelo celular):
 *  1. Crie uma planilha nova no Google Sheets (ex.: "Huddle - Histórico Diário")
 *  2. Copie o ID dela: é o trecho da URL entre "/d/" e "/edit"
 *     Ex.: https://docs.google.com/spreadsheets/d/AQUI_ESTA_O_ID/edit
 *  3. Cole esse ID na constante SHEET_ID logo abaixo, entre aspas
 *  4. Abra script.google.com (pelo navegador do celular, sem precisar
 *     de "modo desktop") → "Novo projeto"
 *  5. Apague o conteúdo padrão e cole este arquivo inteiro (já com o
 *     SHEET_ID preenchido)
 *  6. Toque no ícone de disquete para salvar
 *  7. "Implantar" > "Nova implantação"
 *     - Tipo: "App da Web"
 *     - Executar como: "Eu" (sua conta)
 *     - Quem tem acesso: "Qualquer pessoa"
 *  8. Autorize as permissões quando solicitado
 *  9. Copie a URL do app da Web (termina em /exec) e me envie
 * 10. (Opcional) Defina uma senha na constante SECRET abaixo — use a
 *     MESMA senha no campo "Senha de proteção" do app
 */

var SHEET_ID = "COLE_AQUI_O_ID_DA_PLANILHA"; // obrigatório — veja o Passo 2/3 acima
var SECRET = ""; // opcional — ex.: "sobral2026". Deixe "" para não exigir senha.

var UNITS = ['cm1','cm2','cg','orto_sjq','orto_sjs','onco'];
var UNIT_LABELS = {
  cm1: 'CM1',
  cm2: 'CM2',
  cg: 'CG',
  orto_sjq: 'Orto São Joaquim',
  orto_sjs: 'Orto São José',
  onco: 'Oncologia'
};

function doPost(e) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Histórico') || ss.insertSheet('Histórico');

  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ ok: false, error: 'payload inválido' });
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
  var headers = ['Carimbo de envio', 'Data', 'Hora', 'Turno', 'Preenchido por', 'CRM-CE', 'Função'];
  UNITS.forEach(function (k) {
    var L = UNIT_LABELS[k];
    headers.push(
      L + ' Pacientes', L + ' Vagos Masc', L + ' Vagos Fem', L + ' Admissões', L + ' Altas',
      L + ' Risco (nº)', L + ' Risco (lista)',
      L + ' Paliativos (nº)', L + ' Paliativos (leitos)',
      L + ' Hemodiálise (nº)', L + ' Hemodiálise (leitos)',
      L + ' Pendências (nº)', L + ' Pendências (lista)',
      L + ' Transferências'
    );
  });
  headers.push(
    'Total Pacientes', 'Total Vagos Masc', 'Total Vagos Fem', 'Total Vagos',
    'Total Admissões', 'Total Altas', 'Total Risco', 'Total Paliativos', 'Total Hemodiálise',
    'Total Pendências', 'Ocupação (%)'
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
        cargo: 'Médico Hospitalista',
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
