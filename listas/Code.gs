/**
 * Listas de Chequeo — backend en Google Apps Script + Google Sheets
 * Usa una Google Sheet PROPIA (distinta a químicos y desperdicio).
 *
 * INSTALACIÓN:
 * 1. Crea una Google Sheet nueva.
 * 2. Extensiones -> Apps Script, borra todo y pega este archivo.
 * 3. Implementar -> Nueva implementación -> Aplicación web
 *    (Ejecutar como: Yo · Acceso: Cualquier usuario).
 * 4. Copia la URL /exec y pégala en "Sincronización en tiempo real" de la app.
 *
 * Guarda en dos hojas:
 *  - "Inspecciones": una fila por inspección (cabecera + decisión).
 *  - "Detalle": una fila por variable inspeccionada, ligada por inspeccionId.
 */

const SH_INSP = "Inspecciones";
const SH_DET  = "Detalle";
const TZ = "America/Bogota";

const H_INSP = ["id","numero","esReinspeccion","fecha","hora","fechaRecepcion","orden","proveedor","trazabilidad",
  "listaId","producto","codigoBPCS","codigoDoc","version",
  "cantidadRecibida","cantidadMuestra","nivelInspeccion","aql",
  "inspector","lider","decision","observaciones","totalVariables","noConformes"];

const H_DET = ["inspeccionId","orden","producto","fecha","variable","especificacion",
  "valorMedido","resultado","observacion"];

function getSheet_(nombre, headers){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nombre);
  if(!sh){
    sh = ss.insertSheet(nombre);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function doGet(e){
  const sh = getSheet_(SH_INSP, H_INSP);
  const shD = getSheet_(SH_DET, H_DET);
  const data = sh.getDataRange().getValues();
  if(data.length <= 1) return json_([]);

  const headers = data[0];
  const det = shD.getDataRange().getValues();
  const detHead = det.length ? det[0] : H_DET;

  const porInsp = {};
  for(let i=1;i<det.length;i++){
    const o = {};
    detHead.forEach((h,k)=> o[h] = fmt_(det[i][k], h));
    (porInsp[o.inspeccionId] = porInsp[o.inspeccionId] || []).push({
      nombre:o.variable, especificacion:o.especificacion,
      valorMedido:o.valorMedido, resultado:o.resultado, observacion:o.observacion
    });
  }

  const rows = data.slice(1).map(r=>{
    const o = {};
    headers.forEach((h,i)=> o[h] = fmt_(r[i], h));
    o.resultados = porInsp[o.id] || [];
    return o;
  });
  rows.reverse();
  return json_(rows);
}

function fmt_(v, h){
  if(v instanceof Date){
    if(h === "fecha" || h === "fechaRecepcion") return Utilities.formatDate(v, TZ, "yyyy-MM-dd");
    if(h === "hora") return Utilities.formatDate(v, TZ, "HH:mm");
    return Utilities.formatDate(v, TZ, "yyyy-MM-dd HH:mm:ss");
  }
  return v;
}

function doPost(e){
  try{
    const p = JSON.parse(e.postData.contents);
    if(p.action === "create"){ guardar_(p.registro); return json_({ok:true}); }
    if(p.action === "createBatch"){ (p.registros||[]).forEach(guardar_); return json_({ok:true}); }
    return json_({ok:false, error:"acción no reconocida"});
  }catch(err){
    return json_({ok:false, error:err.message});
  }
}

function guardar_(r){
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try{
    const sh = getSheet_(SH_INSP, H_INSP);
    const last = sh.getLastRow();
    if(last > 1){
      const ids = sh.getRange(2,1,last-1,1).getValues().flat();
      if(ids.indexOf(r.id) !== -1) return; // idempotencia
    }
    const nc = (r.resultados||[]).filter(x => x.resultado === "Rechazo").length;
    sh.appendRow(H_INSP.map(h=>{
      if(h === "totalVariables") return (r.resultados||[]).length;
      if(h === "noConformes") return nc;
      return r[h] !== undefined ? r[h] : "";
    }));

    const shD = getSheet_(SH_DET, H_DET);
    (r.resultados||[]).forEach(v=>{
      shD.appendRow([r.id, r.orden, r.producto, r.fecha,
        v.nombre, v.especificacion, v.valorMedido, v.resultado, v.observacion]);
    });
  } finally {
    lock.releaseLock();
  }
}

function json_(o){
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
