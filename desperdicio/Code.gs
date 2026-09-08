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
  "id", "fecha", "hora", "familia", "proceso", "puesto",
  "defectoCodigo", "defectoLabel", "origenDefecto", "defectoAjeno",
  "cantidad", "observacion", "turno", "turnoKey"
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

const SHEET_RAT = "Ratios";
const H_RAT = ["familia", "maquina", "ratio", "activa", "actualizado"];

function getSheetRat_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_RAT);
  if (!sh) { sh = ss.insertSheet(SHEET_RAT); sh.appendRow(H_RAT); sh.setFrozenRows(1); }
  return sh;
}

const SHEET_EMP = "Empacadas";
const H_EMP = ["familia", "clave", "unidades", "actualizado"];

function getSheetEmp_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_EMP);
  if (!sh) { sh = ss.insertSheet(SHEET_EMP); sh.appendRow(H_EMP); sh.setFrozenRows(1); }
  return sh;
}

function doGet(e) {
  if (e && e.parameter && e.parameter.tipo === "ratios") {
    const sh = getSheetRat_();
    const d = sh.getDataRange().getValues();
    if (d.length <= 1) return jsonOutput_([]);
    return jsonOutput_(d.slice(1).map(r => ({familia: r[0], maquina: r[1], ratio: r[2], activa: r[3]})));
  }
  if (e && e.parameter && e.parameter.tipo === "empacadas") {
    const sh = getSheetEmp_();
    const d = sh.getDataRange().getValues();
    if (d.length <= 1) return jsonOutput_([]);
    return jsonOutput_(d.slice(1).map(r => ({familia: r[0], clave: r[1], unidades: r[2]})));
  }
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
    if (payload.action === "ratio") {
      const sh = getSheetRat_();
      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        const d = sh.getDataRange().getValues();
        const stamp = Utilities.formatDate(new Date(), "America/Bogota", "yyyy-MM-dd HH:mm:ss");
        for (let i = 1; i < d.length; i++) {
          if (d[i][0] === payload.familia && d[i][1] === payload.maquina) {
            sh.getRange(i + 1, 3, 1, 3).setValues([[payload.ratio, payload.activa, stamp]]);
            return jsonOutput_({ ok: true, actualizado: true });
          }
        }
        sh.appendRow([payload.familia, payload.maquina, payload.ratio, payload.activa, stamp]);
      } finally { lock.releaseLock(); }
      return jsonOutput_({ ok: true });
    }

    if (payload.action === "empacadas") {
      const sh = getSheetEmp_();
      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        const d = sh.getDataRange().getValues();
        const stamp = Utilities.formatDate(new Date(), "America/Bogota", "yyyy-MM-dd HH:mm:ss");
        for (let i = 1; i < d.length; i++) {
          if (d[i][0] === payload.familia && d[i][1] === payload.clave) {
            sh.getRange(i + 1, 3).setValue(payload.unidades);
            sh.getRange(i + 1, 4).setValue(stamp);
            return jsonOutput_({ ok: true, actualizado: true });
          }
        }
        sh.appendRow([payload.familia, payload.clave, payload.unidades, stamp]);
      } finally { lock.releaseLock(); }
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
