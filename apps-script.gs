// ============================================================
// Clarios IS — Google Apps Script Backend
// Despliega: Ejecutar como "Yo", Acceso "Cualquiera"
// ============================================================

const SPREADSHEET_ID = 'TU_SPREADSHEET_ID_AQUI'; // <-- Reemplaza con el ID de tu Google Sheet

function doGet(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetTypes = [
      { name: 'Proyectos',    type: 'proyecto'   },
      { name: 'Eventos',      type: 'evento'     },
      { name: 'Activaciones', type: 'activacion' }
    ];
    const allRecords = [];
    sheetTypes.forEach(({ name, type }) => {
      const sheet = ss.getSheetByName(name);
      if (!sheet || sheet.getLastRow() < 2) return;
      const data = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
      const headers = data[0].map(h => h.toString());
      for (let i = 1; i < data.length; i++) {
        if (!data[i][0]) continue;
        const rec = { type };
        headers.forEach((h, j) => { if (h) rec[h] = data[i][j] != null ? data[i][j].toString() : ''; });
        allRecords.push(rec);
      }
    });
    output.setContent(JSON.stringify({ status: 'ok', data: allRecords, total: allRecords.length }));
  } catch (err) {
    output.setContent(JSON.stringify({ status: 'error', message: err.toString() }));
  }
  return output;
}

function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    const payload = JSON.parse(e.postData.contents);
    const type = (payload.type || '').toLowerCase().trim();
    const sheetMap = { proyecto: 'Proyectos', evento: 'Eventos', activacion: 'Activaciones' };
    const sheetName = sheetMap[type];
    if (!sheetName) throw new Error('Tipo inválido: ' + type);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) { crearHojas(); sheet = ss.getSheetByName(sheetName); }

    const prefixMap = { Proyectos: 'PRO', Eventos: 'EVE', Activaciones: 'ACT' };
    const seq = String(sheet.getLastRow()).padStart(4, '0');
    const folio = prefixMap[sheetName] + '-' + seq + '-' + new Date().getFullYear();

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    payload.folio = folio;
    payload.fecha_registro = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const row = headers.map(h => (payload[h] != null ? payload[h] : ''));
    sheet.appendRow(row);

    output.setContent(JSON.stringify({ status: 'ok', folio, message: 'Registro guardado' }));
  } catch (err) {
    output.setContent(JSON.stringify({ status: 'error', message: err.toString() }));
  }
  return output;
}

// Ejecuta esta función UNA VEZ manualmente desde el editor para crear las hojas
function crearHojas() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const common = [
    'folio','fecha_registro','nombre','fecha','responsable','ubicacion',
    'area','beneficiarios','tipo_beneficiario','objetivo','programa','notas',
    'proposito','liderazgo','reputacion','sostenibilidad','cocreacion','gestion','valor',
    'g_ben_directos','g_horas_formacion','g_pct_mujeres'
  ];
  const defs = {
    'Proyectos':    [...common, 'duracion','presupuesto','aliados','indicador','meta','alineacion'],
    'Eventos':      [...common, 'tipo_evento','asistentes','modalidad','media','logistica'],
    'Activaciones': [...common, 'tipo_activacion','canal','alcance','inversion','mensaje']
  };
  Object.entries(defs).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    else if (sheet.getRange(1,1).getValue() === 'folio') {
      Logger.log(name + ': ya tiene headers, omitiendo.');
      return;
    }
    const r = sheet.getRange(1, 1, 1, headers.length);
    r.setValues([headers]);
    r.setFontWeight('bold');
    r.setBackground('#4A1D8B');
    r.setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(3, 220);
    Logger.log('✓ Hoja creada: ' + name + ' (' + headers.length + ' cols)');
  });
  Logger.log('Setup completado.');
}

// Pruebas locales en el editor
function testGet()  { Logger.log(doGet({}).getContent()); }
function testPost() {
  const mock = { type:'proyecto', nombre:'Test', fecha:'2026-05-27', responsable:'Admin',
    ubicacion:'CDMX', area:'Impacto Social', programa:'General', beneficiarios:100,
    tipo_beneficiario:'Jóvenes', objetivo:'Prueba del sistema',
    proposito:4, liderazgo:4, reputacion:3, sostenibilidad:4, cocreacion:3, gestion:4, valor:4,
    g_ben_directos:80, g_horas_formacion:10, g_pct_mujeres:50 };
  Logger.log(doPost({ postData: { contents: JSON.stringify(mock) } }).getContent());
}
