// =====================================================
// TLC - Registro Asistencia Niños v3
// Estructura: Col A = Nombres, Col B+ = Fechas (dd-mm-yyyy)
// =====================================================

const SHEET_ID_NINOS = '1Ne6YxlWJlq8QObF55Dv3oMM26mSG2kbB-Xlg6yh6OJ4';
const HOJA_NINOS = 'Asistencia ninos';

// Nombres correctos desde planilla TLC
const NINOS = ["AGUSTINA", "AMANDA ZALDIVAR", "AMELIA ZALDIVAR", "AYENDRY ARIAS", "BASTIAN RODRIGUEZ", "CONSTANZA SAN MARTIN", "CRISTOBAL DE LA FUENTE", "FLORENCIA BARROS", "ISABELLA FUENTES", "JOAQUIN LERMANDA", "JOAQUIN SOTO", "JOHAN ARIAS", "JORGITO EMANUEL ROJAS ARAYA", "JOSE MARIA BARROS", "JULIANO ROJAS", "JULIETA RIQUELME", "KARLA SAN MARTIN", "LISBETH ROJAS", "MATIAS ARAYA", "NEYSHA MONESTIME", "YOVANA ARIAS"];

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

function guardarAsistencia(data) {
  const fecha     = data.fecha || '';
  const actividad = data.actividad || '';
  const presentes = typeof data.presentes === 'string'
    ? JSON.parse(data.presentes) : (data.presentes || []);

  if (!fecha) return { ok: false, error: 'Fecha requerida' };

  // Convertir YYYY-MM-DD a dd-mm-yyyy
  const p = fecha.split('-');
  const fechaHeader = `${p[2]}-${p[1]}-${p[0]}`;

  const ss    = SpreadsheetApp.openById(SHEET_ID_NINOS);
  let   sheet = ss.getSheetByName(HOJA_NINOS);

  // Crear hoja si no existe
  if (!sheet) {
    sheet = ss.insertSheet(HOJA_NINOS);
    sheet.getRange(1, 1).setValue('NOMBRE')
      .setFontWeight('bold').setBackground('#1a3a5c')
      .setFontColor('white').setHorizontalAlignment('center');
    NINOS.forEach((n, i) => sheet.getRange(i + 2, 1).setValue(n));
    sheet.setColumnWidth(1, 220);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);
  }

  // Asegurar que todos los niños están en col A
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const nombresExist = sheet.getRange(2, 1, Math.max(lastRow - 1, 1), 1)
    .getValues().flat().map(String);
  NINOS.forEach(n => {
    if (!nombresExist.includes(n)) {
      sheet.getRange(sheet.getLastRow() + 1, 1).setValue(n);
      nombresExist.push(n);
    }
  });

  // Buscar o crear columna para esta fecha
  const lastCol   = Math.max(sheet.getLastColumn(), 1);
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let   colNum    = headerRow.indexOf(fechaHeader) + 1;

  if (colNum === 0) {
    // Nueva columna
    colNum = lastCol + 1;
    sheet.getRange(1, colNum)
      .setValue(fechaHeader)
      .setFontWeight('bold').setBackground('#1a3a5c')
      .setFontColor('white').setHorizontalAlignment('center')
      .setNote(actividad);
    sheet.setColumnWidth(colNum, 110);
  }

  // Limpiar columna
  const totalRows = sheet.getLastRow();
  if (totalRows > 1) {
    const rng = sheet.getRange(2, colNum, totalRows - 1, 1);
    rng.clearContent();
    rng.setBackground(null);
    rng.setFontColor('#374151');
    rng.setHorizontalAlignment('center');
  }

  // Leer nombres actuales
  const nombresCol = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
    .getValues().flat().map(String);

  // Marcar todos como ausente primero
  nombresCol.forEach((n, i) => {
    sheet.getRange(i + 2, colNum)
      .setValue('—').setFontColor('#d1d5db').setHorizontalAlignment('center');
  });

  // Marcar presentes
  presentes.forEach(nombre => {
    const idx = nombresCol.indexOf(nombre);
    if (idx !== -1) {
      sheet.getRange(idx + 2, colNum)
        .setValue('✓').setFontWeight('bold')
        .setFontColor('#15803d').setBackground('#f0fdf4')
        .setHorizontalAlignment('center');
    }
  });

  return { ok: true, total: presentes.length, fecha: fechaHeader };
}

function getRegistros() {
  const ss    = SpreadsheetApp.openById(SHEET_ID_NINOS);
  const sheet = ss.getSheetByName(HOJA_NINOS);
  if (!sheet || sheet.getLastColumn() < 2) return { ok: true, registros: [] };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const registros = [];

  for (let col = 1; col < headers.length; col++) {
    const fechaH = String(headers[col]).trim();
    if (!fechaH || fechaH === 'NOMBRE') continue;
    const presentes = [];
    for (let row = 1; row < data.length; row++) {
      const nombre = String(data[row][0] || '').trim();
      const val    = String(data[row][col] || '').trim();
      if (nombre && val === '✓') presentes.push(nombre);
    }
    const pts = fechaH.split('-');
    const fechaISO = pts.length === 3 ? `${pts[2]}-${pts[1]}-${pts[0]}` : fechaH;
    const note = sheet.getRange(1, col + 1).getNote() || 'CN';
    registros.push({
      id: fechaISO, fecha: fechaISO, fechaDisplay: fechaH,
      actividad: note, presentes, total: presentes.length
    });
  }

  return { ok: true, registros: registros.reverse() };
}
