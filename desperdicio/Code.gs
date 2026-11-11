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
/* Columnas nuevas desde 2026-11-05:
   - tipo .......... "recoleccion" | "auditoria". Las dos cuentan LAS MISMAS
                     piezas, así que jamás se suman. Las filas viejas vienen
                     en blanco y la app las lee como auditoría, que es lo que
                     eran: se capturaron con máquina obligatoria.
   - referencia .... producto al que se le carga el desperdicio.
   - pasada ........ 1 o 2, cuántas veces recorrió calidad ese proceso.
   - usuario ....... quién firmó la auditoría (la recolección va sin firma).
   - estadoMaquina . "Sin desperdicio" o "No ha trabajado" en las filas de
                     cierre, que llevan cantidad 0. Distinguir las dos
                     importa: la primera entra al reparto, la segunda queda
                     fuera. Las filas anteriores al 10-11-2026 dicen
                     "No trabajó"; la app lee las dos formas.
   - escritoEl .... instante en que ESTA HOJA guardó la fila, puesto por el
                     Apps Script y no por la tablet. Es el cursor de las
                     lecturas incrementales. Tiene que ponerlo el servidor:
                     con el reloj de cada tablet, uno atrasado haría que sus
                     filas quedaran "en el pasado" y los demás no las vieran
                     nunca.
   - capturadoEl ... el día en que alguien DIGITÓ la fila, que no es el mismo
                     que `fecha` cuando se carga al día siguiente. `fecha` es
                     la jornada a la que pertenece el desperdicio. Mezclarlos
                     impediría saber qué se apuntó tarde, y es además lo que
                     permite bloquear una jornada ya reportada sin bloquear
                     la que se está cargando esta mañana. */
const HEADERS = [
  "id", "fecha", "hora", "familia", "proceso", "puesto",
  "defectoCodigo", "defectoLabel", "origenDefecto", "defectoAjeno",
  "cantidad", "observacion", "turno", "turnoKey",
  "tipo", "referencia", "pasada", "usuario", "estadoMaquina", "capturadoEl",
  "escritoEl"
];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    return sheet;
  }
  /* La hoja ya existía con 14 columnas. Añadir nombres a HEADERS no basta:
     hay que escribirlos también en la fila 1, o las filas nuevas se guardarían
     en columnas sin título y la lectura las devolvería con la llave vacía.
     Se añaden sólo los que faltan y al final, para no mover lo ya escrito. */
  const ancho = sheet.getLastColumn();
  const actuales = ancho ? sheet.getRange(1, 1, 1, ancho).getValues()[0].map(String) : [];
  const faltan = HEADERS.filter(function (h) { return actuales.indexOf(h) === -1; });
  if (faltan.length) {
    sheet.getRange(1, ancho + 1, 1, faltan.length).setValues([faltan]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

const SHEET_USR = "Usuarios";
const H_USR = ["usuario", "nombre", "pin", "rol", "activo"];

function getSheetUsr_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_USR);
  if (!sh) { sh = ss.insertSheet(SHEET_USR); sh.appendRow(H_USR); sh.setFrozenRows(1); }
  return sh;
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

/* Marca de versión del backend, para saber desde la app si la hoja está
   corriendo este código o uno anterior. */
const BACKEND_VERSION = "2026-11-11-incremental";

function doGet(e) {
  // ninguna excepción debe salir como página HTML: la app la ve como respuesta
  // ilegible y no puede decir qué pasó
  try { return doGet_(e); }
  catch (err) { return jsonOutput_({ error: String(err && err.message ? err.message : err) }); }
}

function doGet_(e) {
  if (e && e.parameter && e.parameter.tipo === "version") {
    return jsonOutput_({ backend: BACKEND_VERSION });
  }

  if (e && e.parameter && e.parameter.tipo === "ratios") {
    const sh = getSheetRat_();
    const d = sh.getDataRange().getValues();
    if (d.length <= 1) return jsonOutput_([]);
    return jsonOutput_(d.slice(1).map(r => ({familia: r[0], maquina: r[1], ratio: r[2], activa: r[3]})));
  }
  /* Usuarios que pueden firmar una auditoría. La recolección no pasa por aquí:
     el patinador registra sin entrar. Si la hoja está vacía la app no bloquea
     la auditoría, deja escribir el nombre y lo advierte. */
  if (e && e.parameter && e.parameter.tipo === "usuarios") {
    const sh = getSheetUsr_();
    const d = sh.getDataRange().getValues();
    if (d.length <= 1) return jsonOutput_([]);
    return jsonOutput_(d.slice(1)
      .filter(function (r) { return String(r[0] || "").trim() !== ""; })
      .map(function (r) {
        return {usuario: String(r[0]), nombre: String(r[1] || ""),
                pin: String(r[2] || ""), rol: String(r[3] || ""),
                activo: String(r[4] === "" || r[4] === undefined ? "Sí" : r[4])};
      }));
  }
  if (e && e.parameter && e.parameter.tipo === "empacadas") {
    const sh = getSheetEmp_();
    const d = sh.getDataRange().getValues();
    if (d.length <= 1) return jsonOutput_([]);
    return jsonOutput_(d.slice(1).map(r => ({familia: r[0], clave: r[1], unidades: r[2]})));
  }
  const sheet = getSheet_();
  const ultima = sheet.getLastRow();
  if (ultima <= 1) return jsonOutput_([]);

  const TZ = "America/Bogota";
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const iFecha = headers.indexOf("fecha");
  const desde = (e && e.parameter && e.parameter.desde) ? String(e.parameter.desde) : "";

  function fechaTxt_(v) {
    return v instanceof Date ? Utilities.formatDate(v, TZ, "yyyy-MM-dd") : String(v || "").slice(0, 10);
  }

  /* Antes se cargaba TODA la hoja en memoria en cada lectura, aunque luego se
     filtrara. Con el histórico creciendo, cada petición tarda más y al coincidir
     varios equipos Google corta la ejecución.
     Ahora se mira sólo la columna de fecha —una columna es barata— para acotar
     el tramo, y se traen únicamente esas filas. No se supone que estén
     ordenadas: un equipo que sincroniza tarde deja una fila vieja al final. */
  /* ---- Lectura incremental ----
     Con `desdeEscrito` sólo se devuelve lo que la hoja guardó DESPUÉS de ese
     instante. Es lo que evita que cada sondeo baje el histórico entero: pasa
     de miles de filas a las pocas que hayan entrado desde la última vez.
     Para localizar el punto de corte se lee una sola columna, que es barato,
     y después sólo el tramo final.
     Lo que esto NO puede ver es lo que alguien BORRE: una fila que desaparece
     no deja marca. Por eso el cliente sigue haciendo de vez en cuando una
     lectura completa; si se confiara sólo en lo incremental, lo borrado desde
     otro equipo seguiría contando aquí para siempre. */
  const desdeEscrito = (e && e.parameter && e.parameter.desdeEscrito)
    ? String(e.parameter.desdeEscrito) : "";
  if (desdeEscrito) {
    const iEsc = headers.indexOf("escritoEl");
    if (iEsc === -1) {
      /* Hoja sin la columna todavía: se avisa para que el cliente vuelva a la
         lectura completa en vez de creerse que no hay nada nuevo. */
      return jsonOutput_({ incremental: false, registros: [], cursor: "" });
    }
    const escritos = sheet.getRange(2, iEsc + 1, ultima - 1, 1).getValues();
    let min = -1, maxEsc = desdeEscrito;
    for (let i = 0; i < escritos.length; i++) {
      const v = escritos[i][0] instanceof Date
        ? Utilities.formatDate(escritos[i][0], TZ, "yyyy-MM-dd HH:mm:ss")
        : String(escritos[i][0] || "");
      if (v && v > desdeEscrito) { if (min === -1) min = i; if (v > maxEsc) maxEsc = v; }
    }
    if (min === -1) return jsonOutput_({ incremental: true, registros: [], cursor: desdeEscrito });
    /* Desde la primera fila nueva hasta el final: lo normal es que las nuevas
       estén todas juntas al final, porque se añaden al final. */
    const tramo = sheet.getRange(min + 2, 1, ultima - (min + 2) + 1, headers.length).getValues();
    const nuevas = [];
    tramo.forEach(function (row) {
      const v = row[iEsc] instanceof Date
        ? Utilities.formatDate(row[iEsc], TZ, "yyyy-MM-dd HH:mm:ss")
        : String(row[iEsc] || "");
      if (!v || v <= desdeEscrito) return;       // por si el tramo trae alguna vieja
      nuevas.push(filaAObjeto_(headers, row, TZ));
    });
    return jsonOutput_({ incremental: true, registros: nuevas, cursor: maxEsc });
  }

  let primera = 2, ultimaFila = ultima;
  if (desde && iFecha !== -1) {
    const fechas = sheet.getRange(2, iFecha + 1, ultima - 1, 1).getValues();
    let min = -1, max = -1;
    for (let i = 0; i < fechas.length; i++) {
      if (fechaTxt_(fechas[i][0]) >= desde) { if (min === -1) min = i; max = i; }
    }
    if (min === -1) return jsonOutput_([]);
    primera = min + 2; ultimaFila = max + 2;
  }

  const filas = sheet.getRange(primera, 1, ultimaFila - primera + 1, headers.length).getValues();

  const rows = [];
  filas.forEach(function (row) {
    if (desde && iFecha !== -1 && fechaTxt_(row[iFecha]) < desde) return;
    rows.push(filaAObjeto_(headers, row, TZ));
  });
  rows.reverse();
  /* La lectura completa devuelve también el cursor, para que el cliente pueda
     seguir con lecturas incrementales a partir de aquí. Se manda aparte, en
     una cabecera del propio JSON, sólo cuando lo piden: así los clientes
     viejos siguen recibiendo la lista pelada que esperan. */
  if (e && e.parameter && e.parameter.conCursor === "1") {
    const iEsc = headers.indexOf("escritoEl");
    let cursor = "";
    if (iEsc !== -1 && ultima > 1) {
      const escritos = sheet.getRange(2, iEsc + 1, ultima - 1, 1).getValues();
      escritos.forEach(function (c) {
        const v = c[0] instanceof Date
          ? Utilities.formatDate(c[0], TZ, "yyyy-MM-dd HH:mm:ss") : String(c[0] || "");
        if (v > cursor) cursor = v;
      });
    }
    return jsonOutput_({ incremental: false, registros: rows, cursor: cursor });
  }
  return jsonOutput_(rows);
}

/* Una fila de la hoja a objeto, con las fechas en texto para que no viajen
   como marcas de tiempo que cada navegador interprete a su manera. */
function filaAObjeto_(headers, row, TZ) {
  const obj = {};
  headers.forEach(function (h, i) {
    let v = row[i];
    if (v instanceof Date) {
      if (h === "fecha") v = Utilities.formatDate(v, TZ, "yyyy-MM-dd");
      else if (h === "hora") v = Utilities.formatDate(v, TZ, "HH:mm");
      else v = Utilities.formatDate(v, TZ, "yyyy-MM-dd HH:mm:ss");
    }
    obj[h] = v;
  });
  return obj;
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet = getSheet_();

    if (payload.action === "create") {
      appendRegistro_(sheet, payload.registro);
      return jsonOutput_({ ok: true });
    }

    /* Borrar filas mal cargadas, por id.
       Es un borrado REAL: la fila desaparece de la hoja y no queda rastro de
       que existió. Se eligió así a propósito frente a anularla con una fila
       de corrección. La consecuencia hay que tenerla presente: si alguien ya
       sacó un informe con ese número, nadie podrá explicar después por qué
       cambió.
       Las filas se eliminan de abajo hacia arriba. Borrando de arriba abajo,
       cada eliminación sube una posición todas las de debajo y los índices
       que quedan por procesar apuntarían a la fila equivocada. */
    if (payload.action === "borrar") {
      const ids = payload.ids || [];
      if (!ids.length) return jsonOutput_({ ok: false, error: "no se indicó qué borrar" });
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        const ultima = sheet.getLastRow();
        if (ultima < 2) return jsonOutput_({ ok: true, borradas: 0 });
        /* La columna del id se busca por su TÍTULO, no se da por hecho que
           sea la primera. Al escribir ya se respeta el orden real de la fila
           1, así que si alguien movió una columna a mano y aquí fuéramos por
           posición fija, se compararían ids contra otra cosa: no borraría lo
           pedido, o borraría una fila que no era. */
        const titulos = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
        const cId = titulos.indexOf("id");
        if (cId === -1) return jsonOutput_({ ok: false, error: "la hoja no tiene columna 'id'" });
        const col = sheet.getRange(2, cId + 1, ultima - 1, 1).getValues();
        const filas = [];
        for (let i = 0; i < col.length; i++) {
          if (ids.indexOf(String(col[i][0])) !== -1) filas.push(i + 2);
        }
        for (let i = filas.length - 1; i >= 0; i--) sheet.deleteRow(filas[i]);
        SpreadsheetApp.flush();
        return jsonOutput_({ ok: true, borradas: filas.length, pedidas: ids.length });
      } finally { lock.releaseLock(); }
    }
    /* Alta o cambio de un usuario. Se busca por el identificador, así que
       mandar el mismo dos veces actualiza en vez de duplicar. Borrar no se
       ofrece: un usuario que ya firmó auditorías tiene que seguir existiendo
       para que esas firmas signifiquen algo. Se desactiva con activo="No". */
    if (payload.action === "usuario") {
      const sh = getSheetUsr_();
      const u = payload.usuario || {};
      const id = String(u.usuario || "").trim();
      if (!id) return jsonOutput_({ ok: false, error: "falta el identificador" });
      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        const d = sh.getDataRange().getValues();
        const fila = [id, String(u.nombre || ""), String(u.pin || ""),
                      String(u.rol || ""), String(u.activo || "Sí")];
        for (let i = 1; i < d.length; i++) {
          if (String(d[i][0]).trim() === id) {
            sh.getRange(i + 1, 1, 1, H_USR.length).setValues([fila]);
            return jsonOutput_({ ok: true, actualizado: true });
          }
        }
        sh.appendRow(fila);
      } finally { lock.releaseLock(); }
      return jsonOutput_({ ok: true, creado: true });
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

    /* Varias casillas en una sola petición.
       Antes la app mandaba una petición por casilla: al guardar un mes eran ~93
       peticiones simultáneas peleando por el mismo bloqueo, y las que no
       alcanzaban a entrar se perdían sin avisar. Aquí entra todo junto, se
       resuelve con un solo bloqueo y se responde cuántas quedaron escritas. */
    if (payload.action === "empacadasLote") {
      const sh = getSheetEmp_();
      const lock = LockService.getScriptLock();
      lock.waitLock(30000);
      try {
        const items = payload.items || [];
        const d = sh.getDataRange().getValues();
        const stamp = Utilities.formatDate(new Date(), "America/Bogota", "yyyy-MM-dd HH:mm:ss");

        const donde = {};
        for (let i = 1; i < d.length; i++) donde[d[i][0] + "||" + d[i][1]] = i;

        const nuevas = [];
        let actualizadas = 0;
        items.forEach(function (x) {
          const k = payload.familia + "||" + x.clave;
          if (donde[k] !== undefined && donde[k] > 0) {
            const fila = donde[k];
            d[fila][2] = x.unidades;
            d[fila][3] = stamp;
            actualizadas++;
          } else if (donde[k] === undefined) {
            nuevas.push([payload.familia, x.clave, x.unidades, stamp]);
            donde[k] = -1;   // evita duplicar si la llave viene repetida en el lote
          }
        });

        if (actualizadas && d.length > 1) sh.getRange(2, 1, d.length - 1, 4).setValues(d.slice(1));
        if (nuevas.length) sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, 4).setValues(nuevas);

        return jsonOutput_({ ok: true, actualizadas: actualizadas, nuevas: nuevas.length });
      } finally { lock.releaseLock(); }
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
    /* La marca de escritura la pone la hoja, nunca el cliente: es lo que
       permite a cualquier equipo pedir "lo guardado después de X" y fiarse
       del resultado aunque su propio reloj esté mal. */
    r.escritoEl = Utilities.formatDate(new Date(), "America/Bogota", "yyyy-MM-dd HH:mm:ss");
    /* Se escribe siguiendo el orden REAL de la fila 1, no el de HEADERS. Si
       alguien reordenó o insertó una columna a mano, escribir por HEADERS
       metería cada dato en la columna de al lado sin avisar. */
    const orden = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    sheet.appendRow(orden.map(function (h) {
      return r[h] !== undefined && r[h] !== null ? r[h] : "";
    }));
  } finally {
    lock.releaseLock();
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}


/* ======================================================================
   MANTENIMIENTO — renombrar un proceso en el histórico
   ======================================================================
   Cuando una línea le cambia el nombre a un proceso, las filas ya escritas
   conservan el nombre viejo y dejan de aparecer en las vistas por proceso,
   que recorren la lista de procesos actual. Esto las pone al día.

   NO se expone por doGet ni doPost a propósito: es una operación que
   reescribe histórico y no tiene por qué poder dispararse desde la web.
   Se ejecuta a mano desde el editor de Apps Script.

   CÓMO USARLO
   1. Extensiones -> Apps Script.
   2. Arriba, en la lista de funciones, elige  verRenombreLimasEmpaqueAControl
      y pulsa Ejecutar. NO cambia nada: sólo dice cuántas filas tocaría.
      El resultado sale en "Registro de ejecución".
   3. Si el número cuadra, elige  aplicarRenombreLimasEmpaqueAControl
      y pulsa Ejecutar.
   4. Si algo saliera mal: Archivo -> Historial de versiones en la propia
      hoja devuelve el estado anterior.
   ====================================================================== */

/**
 * Cambia el nombre de un proceso sólo dentro de una familia.
 * Acotar por familia es el punto: "Empaque" existe en Machetes, Cuchillas y
 * Palas, y ahí NO se toca. Un reemplazar-todo de la hoja los arrastraría.
 *
 * Es idempotente: una segunda ejecución no encuentra nada que cambiar.
 *
 * @param {string} familia  Familia afectada, p. ej. "Limas".
 * @param {string} viejo    Nombre actual del proceso.
 * @param {string} nuevo    Nombre nuevo.
 * @param {boolean} aplicar false = sólo contar; true = escribir.
 * @return {Object} resumen con lo contado o lo cambiado.
 */
function renombrarProceso_(familia, viejo, nuevo, aplicar) {
  const sh = getSheet_();
  const ultima = sh.getLastRow();
  if (ultima < 2) {
    Logger.log("La hoja no tiene registros.");
    return { filas: 0, cambiadas: 0 };
  }

  /* Se leen los títulos reales de la fila 1 en vez de confiar en HEADERS:
     si alguien movió una columna a mano, ir por posición fija escribiría
     el nombre del proceso encima de otra cosa. */
  const titulos = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const cFam = titulos.indexOf("familia");
  const cPro = titulos.indexOf("proceso");
  if (cFam === -1 || cPro === -1) {
    throw new Error("No encuentro las columnas 'familia' y 'proceso' en la fila 1.");
  }

  const n = ultima - 1;
  const fams  = sh.getRange(2, cFam + 1, n, 1).getValues();
  const procs = sh.getRange(2, cPro + 1, n, 1).getValues();

  const ejemplos = [];
  let cambiadas = 0;
  for (let i = 0; i < n; i++) {
    if (String(fams[i][0]).trim() !== familia) continue;
    if (String(procs[i][0]).trim() !== viejo) continue;
    cambiadas++;
    if (ejemplos.length < 5) ejemplos.push("fila " + (i + 2));
    procs[i][0] = nuevo;
  }

  Logger.log("Familia %s · «%s» -> «%s»", familia, viejo, nuevo);
  Logger.log("Filas revisadas: %s", n);
  Logger.log("Filas que coinciden: %s", cambiadas);
  if (ejemplos.length) Logger.log("Primeras: %s", ejemplos.join(", "));

  if (!aplicar) {
    Logger.log("SIMULACIÓN — no se escribió nada. Ejecuta la función 'aplicar...' para hacerlo.");
    return { filas: n, cambiadas: cambiadas, aplicado: false };
  }
  if (!cambiadas) {
    Logger.log("No hay nada que cambiar.");
    return { filas: n, cambiadas: 0, aplicado: true };
  }

  /* Se escribe la columna entera de una vez. Celda por celda serían miles de
     llamadas y el script se quedaría sin cuota antes de terminar. */
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    sh.getRange(2, cPro + 1, n, 1).setValues(procs);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  Logger.log("LISTO: %s fila(s) actualizadas.", cambiadas);
  return { filas: n, cambiadas: cambiadas, aplicado: true };
}

/** Paso 1 — sólo cuenta, no escribe. Ejecuta esta primero. */
function verRenombreLimasEmpaqueAControl() {
  return renombrarProceso_("Limas", "Empaque", "Control", false);
}

/** Paso 2 — escribe los cambios. Ejecuta esta cuando el conteo cuadre. */
function aplicarRenombreLimasEmpaqueAControl() {
  return renombrarProceso_("Limas", "Empaque", "Control", true);
}
