/**
 * Control de Desperdicio — backend en Google Apps Script + Google Sheets
 * IMPORTANTE: usa una Google Sheet DISTINTA a la de químicos (datos separados).
 *
 * INSTALACIÓN (igual que químicos):
 * 1. Crea una Google Sheet nueva.
 * 2. Extensiones -> Apps Script, borra el contenido por defecto y pega este archivo.
 * 3. Implementar -> Nueva implementación -> Aplicación web.
 *    - Ejecutar como: Yo
 *    - Quién tiene acceso: Cualquier usuario
 * 4. Copia la URL /exec y pégala en "Sincronización en tiempo real" de la app.
 */

const SHEET_NAME = "Desperdicios";
const HEADERS = [
  "id", "fecha", "hora", "familia", "proceso", "puesto", "referencia",
  "defectoCodigo", "defectoLabel", "cantidad", "observacion", "turno", "turnoKey"
];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doGet(e) {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonOutput_([]);

  const headers = data[0];
  const TZ = "America/Bogota";
  const rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      let v = row[i];
      if (v instanceof Date) {
        if (h === "fecha") v = Utilities.formatDate(v, TZ, "yyyy-MM-dd");
        else if (h === "hora") v = Utilities.formatDate(v, TZ, "HH:mm");
        else v = Utilities.formatDate(v, TZ, "yyyy-MM-dd HH:mm:ss");
      }
      obj[h] = v;
    });
    return obj;
  });
  rows.reverse();
  return jsonOutput_(rows);
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet = getSheet_();

    if (payload.action === "create") {
      appendRegistro_(sheet, payload.registro);
      return jsonOutput_({ ok: true });
    }
    if (payload.action === "createBatch") {
      (payload.registros || []).forEach(r => appendRegistro_(sheet, r));
      return jsonOutput_({ ok: true, count: (payload.registros || []).length });
    }
    return jsonOutput_({ ok: false, error: "acción no reconocida" });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message });
  }
}

function appendRegistro_(sheet, r) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const idsExistentes = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
      if (idsExistentes.indexOf(r.id) !== -1) return; // idempotencia
    }
    sheet.appendRow(HEADERS.map(h => r[h] !== undefined ? r[h] : ""));
  } finally {
    lock.releaseLock();
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
