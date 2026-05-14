// =====================================================
// TLC - Registro Asistencia Niños v2
// Google Apps Script - Planilla Asistencia Ninos
// =====================================================

const SHEET_ID_NINOS = '1Ne6YxlWJlq8QObF55Dv3oMM26mSG2kbB-Xlg6yh6OJ4';
const HOJA_NINOS = 'Asistencia Ninos';

function makeResp(data, cb) {
  const json = JSON.stringify(data);
  if (cb) return ContentService
    .createTextOutput(`${cb}(${json})`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const action = e.parameter.action || '';
  const cb     = e.parameter.callback || '';
  try {
    if (action === 'getRegistros')      return makeResp(getRegistros(), cb);
    if (action === 'guardarAsistencia') return makeResp(guardarAsistencia(e.parameter), cb);
    return makeResp({ error: 'Accion no reconocida: ' + action }, cb);
  } catch(err) {
    return makeResp({ error: err.toString() }, cb);
  }
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData ? e.postData.contents : '{}'); }
  catch(err) { try { body = JSON.parse(e.parameter.payload || '{}'); } catch(e2){} }
  try {
    if (body.action === 'guardarAsistencia') return makeResp(guardarAsistencia(body.data), '');
    return makeResp({ error: 'Accion no reconocida' }, '');
  } catch(err) {
    return makeResp({ error: err.toString() }, '');
  }
}

// ── Leer registros ────────────────────────────────────
function getRegistros() {
  const ss = SpreadsheetApp.openById(SHEET_ID_NINOS);
  let sheet = ss.getSheetByName(HOJA_NINOS);
  if (!sheet) return { ok: true, registros: [] };

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { ok: true, registros: [] };

  const registros = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    try {
      registros.push({
        id:        String(row[0]),
        fecha:     String(row[1]),
        actividad: String(row[2]),
        presentes: JSON.parse(row[3] || '[]'),
        total:     row[4],
        registrado_en: String(row[5] || '')
      });
    } catch(e) {}
  }
  return { ok: true, registros };
}

// ── Guardar asistencia ────────────────────────────────
function guardarAsistencia(data) {
  // Acepta tanto objeto directo como parámetros GET
  const id        = data.id        || data.fecha + '_' + data.actividad;
  const fecha     = data.fecha;
  const actividad = data.actividad;
  const presentes = typeof data.presentes === 'string'
    ? JSON.parse(data.presentes) : (data.presentes || []);

  const ss = SpreadsheetApp.openById(SHEET_ID_NINOS);
  let sheet = ss.getSheetByName(HOJA_NINOS);

  if (!sheet) {
    sheet = ss.insertSheet(HOJA_NINOS);
    const header = ['ID', 'Fecha', 'Actividad', 'Presentes (JSON)', 'Total', 'Registrado En'];
    sheet.appendRow(header);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, header.length)
      .setFontWeight('bold')
      .setBackground('#1a3a5c')
      .setFontColor('white');
  }

  const allData  = sheet.getDataRange().getValues();
  const presJSON = JSON.stringify(presentes);
  const total    = presentes.length;
  const now      = new Date().toISOString();

  // Buscar fila existente
  for (let i = 1; i < allData.length; i++) {
    if (String(allData[i][0]) === String(id)) {
      sheet.getRange(i+1, 1, 1, 6).setValues([[id, fecha, actividad, presJSON, total, now]]);
      marcarNombresCols(sheet, i+1, allData[0], presentes);
      return { ok: true, action: 'updated', total };
    }
  }

  // Nueva fila
  sheet.appendRow([id, fecha, actividad, presJSON, total, now]);
  const newRow = sheet.getLastRow();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  marcarNombresCols(sheet, newRow, headers, presentes);
  return { ok: true, action: 'created', total };
}

function marcarNombresCols(sheet, rowNum, headers, presentes) {
  presentes.forEach(nombre => {
    let colIdx = headers.indexOf(nombre);
    if (colIdx === -1) {
      const newCol = headers.length + 1;
      sheet.getRange(1, newCol).setValue(nombre)
        .setFontWeight('bold').setBackground('#dbeafe').setFontColor('#1d4ed8');
      headers.push(nombre);
      colIdx = headers.length - 1;
    }
    sheet.getRange(rowNum, colIdx + 1).setValue('✓');
  });
}
