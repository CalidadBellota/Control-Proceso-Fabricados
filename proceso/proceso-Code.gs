/**
 * Control de Proceso — backend en Google Apps Script + Google Sheets
 * Usa una Google Sheet PROPIA (distinta a listas, químicos y desperdicio).
 *
 * INSTALACIÓN:
 * 1. Crea una Google Sheet nueva.
 * 2. Extensiones -> Apps Script, borra todo y pega este archivo.
 * 3. Implementar -> Nueva implementación -> Aplicación web
 *    (Ejecutar como: Yo · Acceso: Cualquier usuario).
 * 4. Copia la URL /exec y pégala en "Sincronización en tiempo real" de la app.
 *
 * Hojas:
 *  - "Formatos":  un formato de control por fila (sus variables van en JSON).
 *  - "Registros": una fila por control diligenciado.
 *  - "Detalle":   una fila por variable medida, ligada por registroId.
 *  - "Usuarios":  usuario, nombre, pin, rol.
 */

const SH_FOR = "Formatos";
const SH_REG = "Registros";
const SH_DET = "Detalle";
const SH_USR = "Usuarios";
const TZ = "America/Bogota";

const H_FOR = ["id","linea","proceso","nombre","codigoDoc","condicional","encabezado","maquinas",
               "variables","actualizado"];

const H_REG = ["id","fecha","hora","turno","linea","proceso","formatoId","codigoDoc",
               "ordenFabricacion","referencia","tipoAcero","proveedor","espesorAcero","maquina",
               "inspector","revisadoPor","observaciones",
               "totalVariables","fueraDeRango","criticasFuera"];

const H_DET = ["registroId","fecha","linea","proceso","ordenFabricacion","referencia","maquina",
               "variable","tipo","unidad","critica","limiteInferior","limiteSuperior",
               "valor","resultado"];

const H_USR = ["usuario","nombre","pin","rol","activo","actualizado"];

/** Devuelve la hoja y garantiza que el encabezado tenga todas las columnas. */
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

function headerDe_(sh){
  return sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0]
           .map(function(h){ return String(h).trim(); });
}

function filaSegunHeader_(sh, obj){
  return headerDe_(sh).map(function(h){
    if(h === "") return "";
    const v = obj[h];
    if(v === undefined || v === null) return "";
    return (typeof v === "object") ? JSON.stringify(v) : v;
  });
}

/**
 * Escribe sin que Sheets reinterprete el contenido: una medida como "3-5"
 * o una especificación como "1/2" se convertían en fecha.
 */
function anexarComoTexto_(sh, filas){
  if(!filas.length) return;
  const rango = sh.getRange(sh.getLastRow() + 1, 1, filas.length, filas[0].length);
  rango.setNumberFormat("@");
  rango.setValues(filas.map(function(f){
    return f.map(function(c){ return (c === null || c === undefined) ? "" : String(c); });
  }));
}

function fmt_(v, h){
  if(v instanceof Date){
    if(h === "fecha") return Utilities.formatDate(v, TZ, "yyyy-MM-dd");
    if(h === "hora")  return Utilities.formatDate(v, TZ, "HH:mm");
    return Utilities.formatDate(v, TZ, "yyyy-MM-dd");
  }
  return v;
}

function jsonSeguro_(txt, porDefecto){
  try{ return JSON.parse(txt || porDefecto); }catch(err){ return JSON.parse(porDefecto); }
}

function doGet(e){
  // catálogo compartido: formatos de control y usuarios
  if(e && e.parameter && e.parameter.tipo === "config"){
    const shF = getSheet_(SH_FOR, H_FOR);
    const shU = getSheet_(SH_USR, H_USR);

    const dU = shU.getDataRange().getValues();
    if(dU.length <= 1){
      shU.appendRow(["jefe","Jefe de Calidad","1234","jefe","Sí",
        Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm:ss")]);
    }
    const usuarios = shU.getDataRange().getValues().slice(1).filter(function(r){ return r[0]; })
      .map(function(r){ return {usuario:String(r[0]), nombre:String(r[1]), pin:String(r[2]).trim(),
                                rol:String(r[3]||"inspector"), activo:String(r[4]||"Sí")}; });

    const head = headerDe_(shF);
    const col = function(n){ return head.indexOf(n); };
    const formatos = shF.getDataRange().getValues().slice(1)
      .filter(function(r){ return r[col("id")]; })
      .map(function(r){
        return {
          id:          String(r[col("id")]),
          linea:       String(r[col("linea")]),
          proceso:     String(r[col("proceso")]),
          nombre:      String(r[col("nombre")]),
          codigoDoc:   String(r[col("codigoDoc")]||""),
          condicional: String(r[col("condicional")]||"").toLowerCase().indexOf("s") === 0,
          encabezado:  jsonSeguro_(r[col("encabezado")], "[]"),
          maquinas:    jsonSeguro_(r[col("maquinas")], "[]"),
          variables:   jsonSeguro_(r[col("variables")], "[]")
        };
      });

    return json_({formatos:formatos, usuarios:usuarios});
  }

  // registros
  const sh  = getSheet_(SH_REG, H_REG);
  const shD = getSheet_(SH_DET, H_DET);
  const data = sh.getDataRange().getValues();
  if(data.length <= 1) return json_([]);

  const headers = data[0];
  const det = shD.getDataRange().getValues();
  const detHead = det.length ? det[0] : H_DET;

  const porReg = {};
  for(let i = 1; i < det.length; i++){
    const o = {};
    detHead.forEach(function(h,k){ o[h] = fmt_(det[i][k], h); });
    (porReg[o.registroId] = porReg[o.registroId] || []).push({
      nombre:o.variable, tipo:o.tipo, unidad:o.unidad, critica:o.critica,
      min:o.limiteInferior, max:o.limiteSuperior, valor:o.valor, resultado:o.resultado
    });
  }

  const rows = data.slice(1).map(function(r){
    const o = {};
    headers.forEach(function(h,i){ o[h] = fmt_(r[i], h); });
    o.resultados = porReg[o.id] || [];
    return o;
  });
  rows.reverse();
  return json_(rows);
}

function doPost(e){
  try{
    const p = JSON.parse(e.postData.contents);

    if(p.action === "create"){      guardar_(p.registro);                    return json_({ok:true}); }
    if(p.action === "createBatch"){ (p.registros||[]).forEach(guardar_);      return json_({ok:true}); }

    if(p.action === "guardarFormato"){ upsert_(SH_FOR, H_FOR, p.formato, "id");  return json_({ok:true}); }
    if(p.action === "borrarFormato"){  borrar_(SH_FOR, H_FOR, p.id);             return json_({ok:true}); }

    if(p.action === "guardarUsuario"){ upsert_(SH_USR, H_USR, p.usuario, "usuario"); return json_({ok:true}); }
    if(p.action === "borrarUsuario"){  borrar_(SH_USR, H_USR, p.usuarioId);          return json_({ok:true}); }

    if(p.action === "borrarRegistro"){ return json_({ok:true, borrados: borrarRegistro_(p.id)}); }

    return json_({ok:false, error:"acción no reconocida"});
  }catch(err){
    return json_({ok:false, error:err.message});
  }
}

function guardar_(r){
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try{
    const sh = getSheet_(SH_REG, H_REG);
    const last = sh.getLastRow();
    if(last > 1){
      const ids = sh.getRange(2,1,last-1,1).getValues().map(function(x){ return String(x[0]); });
      if(ids.indexOf(String(r.id)) !== -1) return;   // idempotencia
    }

    const res = r.resultados || [];
    const datos = {};
    H_REG.forEach(function(h){ datos[h] = (r[h] !== undefined) ? r[h] : ""; });
    datos.totalVariables = res.length;
    datos.fueraDeRango   = res.filter(function(v){ return v.resultado === "Fuera de rango"; }).length;
    datos.criticasFuera  = res.filter(function(v){
      return v.resultado === "Fuera de rango" && String(v.critica||"").toLowerCase().indexOf("s") === 0;
    }).length;
    anexarComoTexto_(sh, [filaSegunHeader_(sh, datos)]);

    const shD = getSheet_(SH_DET, H_DET);
    const filas = res.map(function(v){
      return filaSegunHeader_(shD, {
        registroId: r.id, fecha: r.fecha, linea: r.linea, proceso: r.proceso,
        ordenFabricacion: r.ordenFabricacion, referencia: r.referencia, maquina: r.maquina,
        variable: v.nombre, tipo: v.tipo, unidad: v.unidad, critica: v.critica,
        limiteInferior: v.min, limiteSuperior: v.max, valor: v.valor, resultado: v.resultado
      });
    });
    anexarComoTexto_(shD, filas);
  } finally { lock.releaseLock(); }
}

function borrarRegistro_(id){
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try{
    let n = 0;
    const sh = getSheet_(SH_REG, H_REG);
    const d = sh.getDataRange().getValues();
    for(let i = d.length - 1; i >= 1; i--){
      if(String(d[i][0]) === String(id)){ sh.deleteRow(i+1); n++; }
    }
    const shD = getSheet_(SH_DET, H_DET);
    const dd = shD.getDataRange().getValues();
    for(let i = dd.length - 1; i >= 1; i--){
      if(String(dd[i][0]) === String(id)) shD.deleteRow(i+1);
    }
    return n;
  } finally { lock.releaseLock(); }
}

function upsert_(nombre, headers, obj, llave){
  const sh = getSheet_(nombre, headers);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try{
    const copia = {};
    headers.forEach(function(h){ copia[h] = obj[h]; });
    copia.actualizado = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm:ss");
    const fila = filaSegunHeader_(sh, copia).map(function(c){
      return (c === null || c === undefined) ? "" : String(c);
    });
    const d = sh.getDataRange().getValues();
    for(let i = 1; i < d.length; i++){
      if(String(d[i][0]) === String(obj[llave])){
        const rg = sh.getRange(i+1, 1, 1, fila.length);
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
  for(let i = 1; i < d.length; i++){
    if(String(d[i][0]) === String(valor)){ sh.deleteRow(i+1); return; }
  }
}

function json_(o){
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
