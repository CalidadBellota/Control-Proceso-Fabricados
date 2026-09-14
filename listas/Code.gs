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

const SH_LIS  = "Listas";
const SH_ORD  = "Ordenes";
const H_LIS = ["id","nombre","codigoBPCS","codigoDoc","version","nivelInspeccion","aql","variables","actualizado"];
const H_ORD = ["numero","proveedor","productos","actualizado"];

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
  // catálogo compartido: listas de chequeo y órdenes de compra
  if(e && e.parameter && e.parameter.tipo === "config"){
    const shL = getSheet_(SH_LIS, H_LIS), shO = getSheet_(SH_ORD, H_ORD);
    const dL = shL.getDataRange().getValues(), dO = shO.getDataRange().getValues();
    const listas = dL.slice(1).filter(r=>r[0]).map(r=>({
      id:String(r[0]), nombre:r[1], codigoBPCS:String(r[2]), codigoDoc:r[3], version:String(r[4]),
      nivelInspeccion:String(r[5]), aql:Number(r[6])||10,
      variables: (function(){ try{ return JSON.parse(r[7]||"[]"); }catch(err){ return []; } })()
    }));
    const ordenes = dO.slice(1).filter(r=>r[0]).map(r=>({
      numero:String(r[0]), proveedor:r[1],
      productos: (function(){ try{ return JSON.parse(r[2]||"[]"); }catch(err){ return []; } })()
    }));
    return json_({listas:listas, ordenes:ordenes});
  }
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

    if(p.action === "guardarLista"){   upsert_(SH_LIS, H_LIS, p.lista, "id");        return json_({ok:true}); }
    if(p.action === "guardarOrden"){   upsert_(SH_ORD, H_ORD, p.orden, "numero");    return json_({ok:true}); }
    if(p.action === "borrarLista"){    borrar_(SH_LIS, H_LIS, p.id);                 return json_({ok:true}); }
    if(p.action === "borrarOrden"){    borrar_(SH_ORD, H_ORD, p.numero);             return json_({ok:true}); }
    return json_({ok:false, error:"acción no reconocida"});
  }catch(err){
    return json_({ok:false, error:err.message});
  }
}

function upsert_(nombre, headers, obj, llave){
  const sh = getSheet_(nombre, headers);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try{
    const stamp = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm:ss");
    const fila = headers.map(h=>{
      if(h === "actualizado") return stamp;
      const v = obj[h];
      if(v === undefined || v === null) return "";
      return (typeof v === "object") ? JSON.stringify(v) : v;
    });
    const d = sh.getDataRange().getValues();
    for(let i=1;i<d.length;i++){
      if(String(d[i][0]) === String(obj[llave])){
        sh.getRange(i+1,1,1,headers.length).setValues([fila]);
        return;
      }
    }
    sh.appendRow(fila);
  } finally { lock.releaseLock(); }
}

function borrar_(nombre, headers, valor){
  const sh = getSheet_(nombre, headers);
  const d = sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++){
    if(String(d[i][0]) === String(valor)){ sh.deleteRow(i+1); return; }
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
