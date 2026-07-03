// ============================================================
// Clarios IS — Google Apps Script Backend
// Despliega: Ejecutar como "Yo", Acceso "Cualquiera"
// ============================================================

const SPREADSHEET_ID = 'TU_SPREADSHEET_ID_AQUI'; // <-- Reemplaza con el ID de tu Google Sheet

// Columnas base comunes a los tres tipos de hoja
const BASE_HEADERS = [
  'id', 'timestamp', 'folio', 'fecha_registro', 'type',
  'nombre', 'fecha', 'responsable', 'ubicacion', 'area',
  'beneficiarios', 'tipo_beneficiario', 'objetivo', 'programa', 'notas'
];

// Indicadores generales (iguales en los tres tipos)
const INDICADORES_HEADERS = [
  // Resultado
  'g_escuelas', 'g_horas_formacion', 'g_ben_directos', 'g_ben_indirectos',
  'g_pct_mujeres', 'g_alianzas', 'g_prototipos', 'g_narrativas', 'g_inversion',
  'g_quejas_recibidas', 'g_quejas_atendidas', 'g_grado_participacion',
  // Impacto — porcentajes
  'g_pct_habilidades', 'g_pct_autonomia_mujeres', 'g_pct_becarios_sost',
  'g_alumnos_genero',
  // Impacto — Satisfacción (nivel-count + mediana)
  'g_sat_l1', 'g_sat_l2', 'g_sat_l3', 'g_sat_l4', 'g_sat_l5', 'g_sat_mediana',
  // Impacto — Percepción de CLARIOS (nivel-count + mediana)
  'g_perc_l1', 'g_perc_l2', 'g_perc_l3', 'g_perc_l4', 'g_perc_l5', 'g_perc_mediana',
  // Reputación — Reconocimiento (Sí/No + %)
  'g_rec_si', 'g_rec_no', 'g_reconocimiento',
  // Reputación — NPS (conteos + score calculado)
  'g_nps_pos', 'g_nps_neu', 'g_nps_neg', 'g_nps',
  // Reputación — Narrativas
  'g_narrativas',
  // Reputación — Índice de Confianza (Sí/No por pregunta + índice calculado)
  'g_ic_q1_si', 'g_ic_q1_no',
  'g_ic_q2_si', 'g_ic_q2_no',
  'g_ic_q3_si', 'g_ic_q3_no',
  'g_ic_q4_si', 'g_ic_q4_no',
  'g_indice_confianza',
  // Proceso Social
  'g_rondas_mentoria', 'g_canal_seguimiento',
  // Metadatos de secciones activas
  'indic_activos',
  // Programa: Arranca el Futuro
  'af_pct_perfiles', 'af_pct_intervenidos', 'af_gimnasios', 'af_sesiones', 'af_obs',
  'af_talento_canalizado', 'af_pct_beneficiados_clarios',
  // Programa: Guardianes del Planeta
  'gp_guardianes', 'gp_residuos_kg', 'gp_estaciones', 'gp_tipo_residuos', 'gp_obs',
  // Programa: Infraestructura Resiliente
  'inf_arboles', 'inf_intervenciones', 'inf_espacios_digitales', 'inf_tipo', 'inf_obs',
  // Programa: Activación Social
  'as_voluntarios', 'as_num_eventos', 'as_tipo', 'as_horas_voluntariado', 'as_obs'
];

// Headers de la hoja de Evaluaciones Impact Kit
const IK_EVAL_HEADERS = [
  'id', 'timestamp', 'folio', 'fecha', 'nombre',
  'ik_valor_compartido', 'ik_liderazgo', 'ik_sostenibilidad',
  'ik_proposito', 'ik_gestion', 'ik_cocreacion', 'ik_reputacion',
  'ik_score', 'ik_veredicto', 'ik_status', 'status_fecha'
];

// Headers completos por tipo de hoja
const SHEET_HEADERS = {
  'Proyectos':    [...BASE_HEADERS, 'duracion', 'presupuesto', 'aliados', 'indicador', 'meta', ...INDICADORES_HEADERS],
  'Eventos':      [...BASE_HEADERS, 'tipo_evento', 'asistentes', 'modalidad', 'media', 'logistica', ...INDICADORES_HEADERS],
  'Activaciones': [...BASE_HEADERS, 'tipo_activacion', 'canal', 'alcance', 'inversion', 'mensaje', ...INDICADORES_HEADERS],
  'Evaluaciones': IK_EVAL_HEADERS
};

// ─────────────────────────────────────────────────────────────
// doGet — devuelve todos los registros como JSON
// ─────────────────────────────────────────────────────────────
function doGet(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetTypes = [
      { name: 'Proyectos',    type: 'proyecto'      },
      { name: 'Eventos',      type: 'evento'        },
      { name: 'Activaciones', type: 'activacion'    },
      { name: 'Evaluaciones', type: 'evaluacion_ik' }
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

// ─────────────────────────────────────────────────────────────
// doPost — guarda un registro en la hoja correspondiente
// ─────────────────────────────────────────────────────────────
function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    const payload = JSON.parse(e.postData.contents);
    const type = (payload.type || '').toLowerCase().trim();
    const sheetMap = { proyecto: 'Proyectos', evento: 'Eventos', activacion: 'Activaciones', evaluacion_ik: 'Evaluaciones' };
    const sheetName = sheetMap[type];
    if (!sheetName) throw new Error('Tipo inválido: ' + type);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) { crearHojas(); sheet = ss.getSheetByName(sheetName); }

    // Si faltan columnas en la hoja, las agrega antes de guardar
    asegurarColumnas_(sheet, SHEET_HEADERS[sheetName]);

    // Folio correlativo
    const prefixMap = { Proyectos: 'PRO', Eventos: 'EVE', Activaciones: 'ACT', Evaluaciones: 'IK' };
    const seq = String(sheet.getLastRow()).padStart(4, '0');
    const folio = prefixMap[sheetName] + '-' + seq + '-' + new Date().getFullYear();

    // Lee los headers actuales de la hoja (ya actualizados)
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

// ─────────────────────────────────────────────────────────────
// asegurarColumnas_ — agrega al final las columnas que falten
// ─────────────────────────────────────────────────────────────
function asegurarColumnas_(sheet, expectedHeaders) {
  const lastCol = sheet.getLastColumn();
  const existingHeaders = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h.toString())
    : [];

  const missing = expectedHeaders.filter(h => !existingHeaders.includes(h));
  if (missing.length === 0) return;

  const startCol = lastCol + 1;
  const headerRange = sheet.getRange(1, startCol, 1, missing.length);
  headerRange.setValues([missing]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4A1D8B');
  headerRange.setFontColor('#FFFFFF');
  Logger.log('Columnas agregadas a ' + sheet.getName() + ': ' + missing.join(', '));
}

// ─────────────────────────────────────────────────────────────
// crearHojas — ejecuta UNA VEZ para crear las hojas vacías
// ─────────────────────────────────────────────────────────────
function crearHojas() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Object.entries(SHEET_HEADERS).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    } else if (sheet.getRange(1, 1).getValue() !== '') {
      // La hoja ya existe con datos → solo asegura que tenga todas las columnas
      asegurarColumnas_(sheet, headers);
      Logger.log(name + ': hoja existente actualizada.');
      return;
    }
    // Hoja nueva → escribe headers completos
    const r = sheet.getRange(1, 1, 1, headers.length);
    r.setValues([headers]);
    r.setFontWeight('bold');
    r.setBackground('#4A1D8B');
    r.setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 140);
    sheet.setColumnWidth(6, 220);
    Logger.log('✓ Hoja creada: ' + name + ' (' + headers.length + ' cols)');
  });
  Logger.log('Setup completado.');
}

// ─────────────────────────────────────────────────────────────
// actualizarHeaders — ejecuta UNA VEZ si ya tienes hojas con
// las columnas antiguas y quieres agregar las nuevas sin
// borrar los datos existentes
// ─────────────────────────────────────────────────────────────
function actualizarHeaders() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Object.entries(SHEET_HEADERS).forEach(([name, headers]) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) { Logger.log(name + ': hoja no encontrada, omitiendo.'); return; }
    asegurarColumnas_(sheet, headers);
  });
  Logger.log('Headers actualizados en todas las hojas.');
}

// ─────────────────────────────────────────────────────────────
// Pruebas locales en el editor
// ─────────────────────────────────────────────────────────────
function testGet()  { Logger.log(doGet({}).getContent()); }

function testPost() {
  const mock = {
    id: 'IS-test-001',
    timestamp: new Date().toISOString(),
    type: 'proyecto',
    nombre: 'Test completo',
    fecha: '2026-05-28',
    responsable: 'Admin',
    ubicacion: 'CDMX',
    area: 'Impacto Social',
    beneficiarios: 100,
    tipo_beneficiario: 'Jóvenes (18–29)',
    objetivo: 'Prueba del sistema con todos los campos',
    programa: 'arranca',
    duracion: '1–3 meses',
    presupuesto: 50000,
    aliados: 'ITESM',
    indicador: '% participantes con empleo',
    meta: '60%',
    g_escuelas: 2,
    g_horas_formacion: 40,
    g_ben_directos: 100,
    g_ben_indirectos: 300,
    g_pct_mujeres: 45,
    g_alianzas: 3,
    g_prototipos: 1,
    g_narrativas: 2,
    g_pct_habilidades: 70,
    g_pct_autonomia_mujeres: 50,
    g_pct_becarios_sost: 30,
    g_percepcion_clarios: '4 — Alto',
    g_reconocimiento: '4 — Referente local',
    g_satisfaccion: '5 — Muy satisfecho',
    g_rondas_mentoria: 2,
    g_canal_seguimiento: 'WhatsApp / Telegram',
    af_pct_perfiles: 20,
    af_pct_intervenidos: 80,
    af_gimnasios: 1,
    af_sesiones: 8,
    af_obs: 'Test exitoso',
    notas: 'Prueba de integración completa'
  };
  Logger.log(doPost({ postData: { contents: JSON.stringify(mock) } }).getContent());
}
