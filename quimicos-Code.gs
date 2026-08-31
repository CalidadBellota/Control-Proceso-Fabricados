/**
 * Control de Químicos — backend en Google Apps Script + Google Sheets
 *
 * INSTALACIÓN:
 * 1. Crea una Google Sheet nueva (puede estar vacía, este script crea la hoja "Registros" solo).
 * 2. Extensiones -> Apps Script.
 * 3. Borra el contenido por defecto y pega este archivo completo.
 * 4. Implementar -> Nueva implementación -> tipo "Aplicación web".
 *    - Ejecutar como: Yo
 *    - Quién tiene acceso: Cualquier usuario
 * 5. Copia la URL que termina en /exec y pégala en el modal "Sincronización en tiempo real" de la app.
 */

const SHEET_NAME = "Registros";
const HEADERS = [
  "id", "fecha", "hora", "turno", "inspeccionId", "inspector", "observaciones",
  "titulo", "insumo", "unidad", "valorMin", "valorMax", "valorMedido", "resultado", "cantidadAplicada", "tipoCorreccion"
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
  if (data.length <= 1) {
    return jsonOutput_([]);
  }
  const headers = data[0];
  const TZ = "America/Bogota";
  const rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      let v = row[i];
      if (v instanceof Date) {
        // Sheets a veces auto-formatea texto tipo "16:44" o "2026-08-31" como
        // fecha/hora real; lo devolvemos siempre como texto plano y consistente.
        if (h === "fecha") v = Utilities.formatDate(v, TZ, "yyyy-MM-dd");
        else if (h === "hora") v = Utilities.formatDate(v, TZ, "HH:mm");
        else v = Utilities.formatDate(v, TZ, "yyyy-MM-dd HH:mm:ss");
      }
      obj[h] = v;
    });
    return obj;
  });
  rows.reverse(); // más recientes primero
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

    if (payload.action === "delete") {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
      return jsonOutput_({ ok: true });
    }

    return jsonOutput_({ ok: false, error: "acción no reconocida" });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message });
  }
}

function appendRegistro_(sheet, r) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // espera hasta 10s si otra escritura está en curso
  try {
    // idempotencia: si este id ya se guardó (reintento duplicado por red), no lo repite
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const idsExistentes = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
      if (idsExistentes.indexOf(r.id) !== -1) return;
    }
    sheet.appendRow(HEADERS.map(h => r[h] !== undefined ? r[h] : ""));
  } finally {
    lock.releaseLock();
  }
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
