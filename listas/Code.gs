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

const SH_USR  = "Usuarios";
const H_USR = ["usuario","nombre","pin","rol","activo","actualizado"];

const SH_LIS  = "Listas";
const SH_ORD  = "Ordenes";
const H_LIS = ["id","nombre","codigoBPCS","codigoDoc","version","nivelInspeccion","aql","variables","actualizado"];
const H_ORD = ["numero","proveedor","productos","actualizado"];

const SH_INSP = "Inspecciones";
const SH_DET  = "Detalle";
const TZ = "America/Bogota";

const H_INSP = ["id","numero","esReinspeccion","estado","aprobadoPor","fechaDecision",
  "tipoRechazo","motivoRechazo","fecha","hora","fechaRecepcion","orden","proveedor","trazabilidad",
  "listaId","producto","codigoBPCS","codigoDoc","version",
  "cantidadRecibida","cantidadMuestra","nivelInspeccion","aql",
  "inspector","lider","decision","observaciones","totalVariables","noConformes"];

const H_DET = ["inspeccionId","orden","producto","fecha","variable","especificacion",
  "valorMedido","resultado","observacion"];

/**
 * Devuelve la hoja y garantiza que su encabezado tenga TODAS las columnas esperadas.
 * Si a una hoja vieja le faltan columnas nuevas, se agregan al final en vez de
 * desplazar los datos existentes.
 */
function getSheet_(nombre, headers){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nombre);
  if(!sh){
    sh = ss.insertSheet(nombre);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    return sh;
  }
  if(sh.getLastRow() === 0){
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    return sh;
  }
  const actuales = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0]
                     .map(function(h){ return String(h).trim(); });
  const faltantes = headers.filter(function(h){ return actuales.indexOf(h) === -1; });
  if(faltantes.length){
    sh.getRange(1, actuales.length + 1, 1, faltantes.length).setValues([faltantes]);
  }
  return sh;
}

/** Encabezado real de la hoja, que es el que manda para ubicar cada dato. */
function headerDe_(sh){
  return sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0]
           .map(function(h){ return String(h).trim(); });
}

/** Arma la fila siguiendo el orden real del encabezado, no el del código. */
function filaSegunHeader_(sh, obj){
  return headerDe_(sh).map(function(h){
    if(h === "") return "";
    const v = obj[h];
    if(v === undefined || v === null) return "";
    return (typeof v === "object") ? JSON.stringify(v) : v;
  });
}

function doGet(e){
  // catálogo compartido: listas de chequeo y órdenes de compra
  if(e && e.parameter && e.parameter.tipo === "config"){
    const shL = getSheet_(SH_LIS, H_LIS), shO = getSheet_(SH_ORD, H_ORD);
    const shU = getSheet_(SH_USR, H_USR);
    const dU = shU.getDataRange().getValues();
    if(dU.length <= 1){   // siembra el primer jefe para poder entrar la primera vez
      shU.appendRow(["jefe","Jefe de Calidad","1234","jefe","Sí",
        Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm:ss")]);
    }
    const usuarios = shU.getDataRange().getValues().slice(1).filter(function(r){ return r[0]; })
      .map(function(r){ return {usuario:String(r[0]), nombre:String(r[1]), pin:String(r[2]).trim(),
                                rol:String(r[3]||"inspector"), activo:String(r[4]||"Sí")}; });
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
    return json_({listas:listas, ordenes:ordenes, usuarios:usuarios});
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

const COLS_TEXTO = ["variable","especificacion","valorMedido","resultado","observacion",
                    "producto","proveedor","orden","codigoBPCS","codigoDoc","version"];

function fmt_(v, h){
  if(v instanceof Date){
    if(h === "fecha" || h === "fechaRecepcion") return Utilities.formatDate(v, TZ, "yyyy-MM-dd");
    if(h === "hora") return Utilities.formatDate(v, TZ, "HH:mm");
    if(h === "trazabilidad") return Utilities.formatDate(v, TZ, "yyyy-MM");
    // una celda de texto que Sheets convirtió a fecha: se devuelve corta, sin hora
    if(COLS_TEXTO.indexOf(h) !== -1) return Utilities.formatDate(v, TZ, "yyyy-MM-dd");
    return Utilities.formatDate(v, TZ, "yyyy-MM-dd HH:mm:ss");
  }
  if(h === "trazabilidad" && v) return String(v).slice(0,7);
  return v;
}

/**
 * Escribe filas SIN que Sheets reinterprete el contenido.
 * Una especificación como "1/2" o un valor como "3-5" se volvían fecha al guardarse;
 * forzando el formato de texto en el rango, el dato queda tal cual se midió.
 */
function anexarComoTexto_(sh, filas){
  if(!filas.length) return;
  const fila0 = sh.getLastRow() + 1;
  const ancho = filas[0].length;
  const rango = sh.getRange(fila0, 1, filas.length, ancho);
  rango.setNumberFormat("@");
  rango.setValues(filas.map(function(f){
    return f.map(function(c){ return (c === null || c === undefined) ? "" : String(c); });
  }));
}

function doPost(e){
  try{
    const p = JSON.parse(e.postData.contents);
    if(p.action === "create"){ guardar_(p.registro); return json_({ok:true}); }
    if(p.action === "createBatch"){ (p.registros||[]).forEach(guardar_); return json_({ok:true}); }

    if(p.action === "guardarUsuario"){ upsert_(SH_USR, H_USR, p.usuario, "usuario"); return json_({ok:true}); }
    if(p.action === "borrarUsuario"){  borrar_(SH_USR, H_USR, p.usuarioId);            return json_({ok:true}); }
    if(p.action === "decidir"){        return json_(decidir_(p)); }

    if(p.action === "guardarLista"){   upsert_(SH_LIS, H_LIS, p.lista, "id");        return json_({ok:true}); }
    if(p.action === "guardarOrden"){   upsert_(SH_ORD, H_ORD, p.orden, "numero");    return json_({ok:true}); }
    if(p.action === "borrarLista"){    borrar_(SH_LIS, H_LIS, p.id);                 return json_({ok:true}); }
    if(p.action === "borrarOrden"){
      borrar_(SH_ORD, H_ORD, p.numero);
      if(p.conInspecciones) borrarInspeccionesDeOrden_(p.numero);
      return json_({ok:true});
    }
    return json_({ok:false, error:"acción no reconocida"});
  }catch(err){
    return json_({ok:false, error:err.message});
  }
}

/** Registra la decisión del jefe sobre una inspección ya guardada. */
function decidir_(p){
  const sh = getSheet_(SH_INSP, H_INSP);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try{
    const head = headerDe_(sh);
    const d = sh.getDataRange().getValues();
    const col = function(n){ return head.indexOf(n) + 1; };
    for(let i = 1; i < d.length; i++){
      if(String(d[i][0]) === String(p.id)){
        const stamp = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm:ss");
        const set = function(n, v){ if(col(n) > 0) sh.getRange(i+1, col(n)).setValue(v); };
        set("estado", p.estado);
        set("decision", p.decision);
        set("aprobadoPor", p.aprobadoPor || "");
        set("fechaDecision", stamp);
        set("tipoRechazo", p.tipoRechazo || "");
        set("motivoRechazo", p.motivoRechazo || "");
        set("lider", p.aprobadoPor || "");
        return { ok:true, fechaDecision: stamp };
      }
    }
    return { ok:false, error:"No se encontró la inspección " + p.id };
  } finally { lock.releaseLock(); }
}

/**
 * Borra de "Inspecciones" y "Detalle" todo lo que cuelgue de una orden de compra.
 * Se usa cuando se elimina la orden: sin orden no debe quedar nada pendiente ni aprobado.
 */
function borrarInspeccionesDeOrden_(numero){
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try{
    const objetivo = String(numero).trim();
    let borradas = 0;

    const sh = getSheet_(SH_INSP, H_INSP);
    const head = headerDe_(sh);
    const iOrden = head.indexOf("orden");
    const iId = head.indexOf("id");
    const ids = {};
    if(iOrden !== -1){
      const d = sh.getDataRange().getValues();
      for(let i = d.length - 1; i >= 1; i--){
        if(String(d[i][iOrden]).trim() === objetivo){
          if(iId !== -1) ids[String(d[i][iId])] = true;
          sh.deleteRow(i + 1);
          borradas++;
        }
      }
    }

    const shD = getSheet_(SH_DET, H_DET);
    const headD = headerDe_(shD);
    const iOrdD = headD.indexOf("orden");
    const iIdD  = headD.indexOf("inspeccionId");
    const dd = shD.getDataRange().getValues();
    for(let i = dd.length - 1; i >= 1; i--){
      const porOrden = iOrdD !== -1 && String(dd[i][iOrdD]).trim() === objetivo;
      const porId    = iIdD  !== -1 && ids[String(dd[i][iIdD])];
      if(porOrden || porId) shD.deleteRow(i + 1);
    }
    return borradas;
  } finally { lock.releaseLock(); }
}

/**
 * Utilidad manual: borra las inspecciones cuya orden de compra ya no existe
 * en la hoja "Ordenes". Ejecútala desde el editor si quedaron huérfanas.
 */
function limpiarInspeccionesHuerfanas(){
  const shO = getSheet_(SH_ORD, H_ORD);
  const ords = {};
  shO.getDataRange().getValues().slice(1).forEach(function(r){
    if(r[0]) ords[String(r[0]).trim()] = true;
  });
  const sh = getSheet_(SH_INSP, H_INSP);
  const iOrden = headerDe_(sh).indexOf("orden");
  if(iOrden === -1) return "No se encontró la columna orden";
  const d = sh.getDataRange().getValues();
  const sueltas = [];
  for(let i = 1; i < d.length; i++){
    const o = String(d[i][iOrden]).trim();
    if(o && !ords[o] && sueltas.indexOf(o) === -1) sueltas.push(o);
  }
  let total = 0;
  sueltas.forEach(function(o){ total += borrarInspeccionesDeOrden_(o); });
  return total + " inspección(es) sin orden eliminadas (" + sueltas.join(", ") + ")";
}

function upsert_(nombre, headers, obj, llave){
  const sh = getSheet_(nombre, headers);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try{
    const stamp = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm:ss");
    const copia = {};
    headers.forEach(function(h){ copia[h] = obj[h]; });
    copia.actualizado = stamp;
    const fila = filaSegunHeader_(sh, copia).map(function(c){
      return (c === null || c === undefined) ? "" : String(c);
    });
    const d = sh.getDataRange().getValues();
    for(let i=1;i<d.length;i++){
      if(String(d[i][0]) === String(obj[llave])){
        const rg = sh.getRange(i+1,1,1,fila.length);
        rg.setNumberFormat("@"); rg.setValues([fila]);
        return;
      }
    }
    anexarComoTexto_(sh, [fila]);
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
    const datos = {};
    H_INSP.forEach(function(h){ datos[h] = (r[h] !== undefined) ? r[h] : ""; });
    datos.totalVariables = (r.resultados||[]).length;
    datos.noConformes = nc;
    if(datos.trazabilidad) datos.trazabilidad = String(datos.trazabilidad).slice(0,7);
    anexarComoTexto_(sh, [filaSegunHeader_(sh, datos)]);

    const shD = getSheet_(SH_DET, H_DET);
    const filasDet = (r.resultados||[]).map(function(v){
      return filaSegunHeader_(shD, {
        inspeccionId: r.id, orden: r.orden, producto: r.producto, fecha: r.fecha,
        variable: v.nombre, especificacion: v.especificacion,
        valorMedido: v.valorMedido, resultado: v.resultado, observacion: v.observacion
      });
    });
    anexarComoTexto_(shD, filasDet);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Utilidad manual: borra de "Inspecciones" las filas cuyo encabezado no cuadra
 * (las que quedaron corridas antes de esta corrección).
 * Ejecútala una sola vez desde el editor si tienes filas desplazadas.
 */
function limpiarFilasCorridas(){
  const sh = getSheet_(SH_INSP, H_INSP);
  const head = headerDe_(sh);
  const iFecha = head.indexOf("fecha");
  if(iFecha === -1) return "No se encontró la columna fecha";
  const d = sh.getDataRange().getValues();
  let borradas = 0;
  for(let i = d.length - 1; i >= 1; i--){
    const v = String(d[i][iFecha] || "");
    // una fecha válida se ve como 2026-09-15; si ahí hay un número suelto, la fila está corrida
    if(v && !/\d{4}-\d{2}-\d{2}/.test(v) && !(d[i][iFecha] instanceof Date)){
      sh.deleteRow(i + 1);
      borradas++;
    }
  }
  return borradas + " fila(s) corridas eliminadas";
}

function json_(o){
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
