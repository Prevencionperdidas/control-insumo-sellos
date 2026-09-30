/**
 * CONTROL DE SELLOS (tapón y fleje) — Backend Google Apps Script
 * PPS IKEA CD Chile
 *
 * Instalación:
 * 1. Crear una Google Sheet nueva > Extensiones > Apps Script > pegar este archivo.
 * 2. Ejecutar la función setup() una vez (autorizar permisos).
 * 3. Implementar > Nueva implementación > Aplicación web
 *    Ejecutar como: Yo  |  Acceso: Cualquier persona
 * 4. Copiar la URL /exec en Ajustes de la app.
 * 5. Para cambios futuros: Gestionar implementaciones > editar > "Nueva versión"
 *    (NO "Nueva implementación", eso cambia la URL).
 */

const PIN_ADMIN = '358097';
const AREAS_DEF = ['Recepción', 'Despacho', 'Bodega', 'Operaciones', 'Administración'];

const HOJAS = {
  Productos: ['Codigo', 'Tipo', 'Descripcion', 'Prefijo', 'StockMinimo', 'Activo'],
  Compras:   ['ID', 'Fecha', 'Origen', 'Factura', 'Proveedor', 'Codigo', 'Tipo', 'Cantidad',
              'FolioDesde', 'FolioHasta', 'Neto', 'IVA', 'Total', 'Usuario', 'Registrado', 'Estado', 'MotivoAnulacion'],
  Entregas:  ['ID', 'Fecha', 'Area', 'Responsable', 'Codigo', 'Tipo', 'Cantidad', 'FolioDesde', 'FolioHasta',
              'Observacion', 'Usuario', 'Registrado', 'Estado', 'MotivoAnulacion'],
  Auditoria: ['Fecha', 'Usuario', 'Accion', 'Detalle'],
  Config:    ['Clave', 'Valor']
};
const COLS_TEXTO = ['ID', 'Factura', 'Codigo', 'Prefijo'];

/* ---------------- Entrada web (JSONP) ---------------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try {
    const data = p.payload ? JSON.parse(p.payload) : {};
    out = route(p.action || 'getAll', data);
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  const json = JSON.stringify(out);
  const cb = p.callback;
  if (cb && /^[A-Za-z_$][\w$]{0,60}$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function route(action, d) {
  switch (action) {
    case 'ping':          return { ok: true, ts: new Date().toISOString() };
    case 'getAll':        return { ok: true, data: getAll() };
    case 'addEntrega':    return conLock(() => addEntrega(d));
    case 'addCompra':     return conLock(() => addCompra(d));
    case 'anularEntrega': return conLock(() => anular('Entregas', d));
    case 'anularCompra':  return conLock(() => anular('Compras', d));
    case 'saveProducto':  return conLock(() => saveProducto(d));
    case 'saveConfig':    return conLock(() => saveConfig(d));
    case 'checkPin':      checkPin(d); return { ok: true };
    default: throw new Error('Acción desconocida: ' + action);
  }
}

function conLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------------- Hojas ---------------- */
function tz() { return Session.getScriptTimeZone() || 'America/Santiago'; }
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function hoja(nombre) {
  let sh = ss().getSheetByName(nombre);
  if (!sh) {
    sh = ss().insertSheet(nombre);
    const h = HOJAS[nombre];
    sh.getRange(1, 1, 1, h.length).setValues([h])
      .setFontWeight('bold').setBackground('#0058A3').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  return sh;
}

function setup() {
  Object.keys(HOJAS).forEach(hoja);
  const cfg = leerConfig();
  if (!cfg.Areas) setConfig('Areas', AREAS_DEF.join(','));
  if (!cfg.Responsables) setConfig('Responsables', '');
  if (!cfg.IVA) setConfig('IVA', '19');
  if (leer('Productos').length === 0) {
    agregar('Productos', { Codigo: 'TAP-01', Tipo: 'Tapón', Descripcion: 'Sello tapón', Prefijo: '', StockMinimo: 100, Activo: 'SI' });
    agregar('Productos', { Codigo: 'FLE-01', Tipo: 'Fleje', Descripcion: 'Sello fleje', Prefijo: '', StockMinimo: 200, Activo: 'SI' });
  }
  const def = ss().getSheetByName('Hoja 1') || ss().getSheetByName('Sheet1');
  if (def && ss().getSheets().length > 1 && def.getLastRow() === 0) ss().deleteSheet(def);
}

function fmtVal(v) {
  if (v instanceof Date) {
    const t = v.getHours() || v.getMinutes() || v.getSeconds();
    return Utilities.formatDate(v, tz(), t ? 'yyyy-MM-dd HH:mm:ss' : 'yyyy-MM-dd');
  }
  return v;
}

function leer(nombre) {
  const sh = hoja(nombre);
  const vals = sh.getDataRange().getValues();
  const h = vals.shift() || [];
  return vals
    .map((r, i) => { const o = { _row: i + 2 }; h.forEach((k, j) => o[k] = fmtVal(r[j])); return o; })
    .filter(o => h.some(k => o[k] !== '' && o[k] !== null));
}

function agregar(nombre, obj) {
  const fila = HOJAS[nombre].map(k => {
    const v = obj[k] === undefined || obj[k] === null ? '' : obj[k];
    return (COLS_TEXTO.indexOf(k) >= 0 && v !== '') ? "'" + v : v;
  });
  hoja(nombre).appendRow(fila);
}

function actualizar(nombre, row, cambios) {
  const sh = hoja(nombre);
  const h = HOJAS[nombre];
  Object.keys(cambios).forEach(k => {
    const c = h.indexOf(k);
    if (c >= 0) sh.getRange(row, c + 1).setValue(cambios[k]);
  });
}

function auditar(usuario, accion, detalle) {
  hoja('Auditoria').appendRow([new Date(), usuario || '—', accion, detalle]);
}

function leerConfig() {
  const o = {};
  leer('Config').forEach(r => o[r.Clave] = r.Valor);
  return o;
}
function setConfig(clave, valor) {
  const r = leer('Config').find(x => x.Clave === clave);
  if (r) actualizar('Config', r._row, { Valor: valor });
  else agregar('Config', { Clave: clave, Valor: valor });
}

function limpiar(arr) { return arr.map(o => { const c = Object.assign({}, o); delete c._row; return c; }); }

function getAll() {
  const aud = leer('Auditoria');
  return {
    productos: limpiar(leer('Productos')),
    compras:   limpiar(leer('Compras')),
    entregas:  limpiar(leer('Entregas')),
    auditoria: limpiar(aud.slice(-400).reverse()),
    config:    leerConfig()
  };
}

/* ---------------- Rangos de folios ---------------- */
function fusionar(iv) {
  const s = iv.filter(x => x[1] >= x[0]).map(x => [x[0], x[1]]).sort((a, b) => a[0] - b[0]);
  const out = [];
  s.forEach(x => {
    const u = out[out.length - 1];
    if (u && x[0] <= u[1] + 1) u[1] = Math.max(u[1], x[1]); else out.push(x);
  });
  return out;
}
function restar(A, B) {
  let res = A.map(x => [x[0], x[1]]);
  B.forEach(b => {
    const n = [];
    res.forEach(a => {
      if (b[1] < a[0] || b[0] > a[1]) { n.push(a); return; }
      if (b[0] > a[0]) n.push([a[0], b[0] - 1]);
      if (b[1] < a[1]) n.push([b[1] + 1, a[1]]);
    });
    res = n;
  });
  return res;
}
const vigente = r => r.Estado !== 'Anulada';
const seSolapan = (a1, b1, a2, b2) => a1 <= b2 && a2 <= b1;

function entero(v, campo) {
  const n = Number(String(v).replace(/\D/g, ''));
  if (!String(v).trim() || !Number.isFinite(n)) throw new Error('Falta ' + campo);
  return n;
}
function checkPin(d) { if (String(d.pin || '') !== PIN_ADMIN) throw new Error('PIN incorrecto'); }

/* ---------------- Acciones ---------------- */
function addEntrega(d) {
  if (!d.ID) throw new Error('Falta ID');
  const E = leer('Entregas');
  if (E.some(r => String(r.ID) === String(d.ID))) return { ok: true, duplicado: true };

  const prod = leer('Productos').find(p => String(p.Codigo) === String(d.Codigo));
  if (!prod) throw new Error('El sello ' + d.Codigo + ' no existe');
  if (!d.Area) throw new Error('Falta el área');
  if (!d.Responsable) throw new Error('Falta el responsable');
  const desde = entero(d.FolioDesde, 'folio desde');
  const hasta = entero(d.FolioHasta, 'folio hasta');
  if (hasta < desde) throw new Error('El folio hasta es menor que el folio desde');

  const cod = String(d.Codigo);
  const recibidos = fusionar(leer('Compras').filter(r => vigente(r) && String(r.Codigo) === cod)
    .map(r => [Number(r.FolioDesde), Number(r.FolioHasta)]));
  const entVig = E.filter(r => vigente(r) && String(r.Codigo) === cod);

  const choque = entVig.find(r => seSolapan(desde, hasta, Number(r.FolioDesde), Number(r.FolioHasta)));
  if (choque) {
    throw new Error('Folios ' + Math.max(desde, choque.FolioDesde) + '–' + Math.min(hasta, choque.FolioHasta) +
      ' ya fueron entregados el ' + choque.Fecha + ' a ' + choque.Responsable + ' (' + choque.Area + ')');
  }
  if (!recibidos.some(x => desde >= x[0] && hasta <= x[1])) {
    throw new Error('Folios ' + desde + '–' + hasta + ' no están registrados como ingresados para ' + cod);
  }

  agregar('Entregas', {
    ID: d.ID, Fecha: d.Fecha, Area: d.Area, Responsable: d.Responsable, Codigo: cod, Tipo: prod.Tipo,
    Cantidad: hasta - desde + 1, FolioDesde: desde, FolioHasta: hasta, Observacion: d.Observacion || '',
    Usuario: d.Usuario || '', Registrado: new Date(), Estado: 'Vigente', MotivoAnulacion: ''
  });
  auditar(d.Usuario, 'Entrega', cod + ' folios ' + desde + '–' + hasta + ' (' + (hasta - desde + 1) + ') a ' +
    d.Responsable + ', ' + d.Area);
  return { ok: true };
}

function addCompra(d) {
  if (!d.ID) throw new Error('Falta ID');
  const C = leer('Compras');
  if (C.some(r => String(r.ID) === String(d.ID))) return { ok: true, duplicado: true };

  const prod = leer('Productos').find(p => String(p.Codigo) === String(d.Codigo));
  if (!prod) throw new Error('El sello ' + d.Codigo + ' no existe');
  const desde = entero(d.FolioDesde, 'folio desde');
  const hasta = entero(d.FolioHasta, 'folio hasta');
  if (hasta < desde) throw new Error('El folio hasta es menor que el folio desde');
  const cod = String(d.Codigo);

  const choque = C.find(r => vigente(r) && String(r.Codigo) === cod &&
    seSolapan(desde, hasta, Number(r.FolioDesde), Number(r.FolioHasta)));
  if (choque) {
    throw new Error('Folios ' + Math.max(desde, choque.FolioDesde) + '–' + Math.min(hasta, choque.FolioHasta) +
      ' ya fueron ingresados el ' + choque.Fecha + (choque.Factura ? ' (factura ' + choque.Factura + ')' : ''));
  }
  const neto = Number(d.Neto) || 0, iva = Number(d.IVA) || 0;
  agregar('Compras', {
    ID: d.ID, Fecha: d.Fecha, Origen: d.Origen || 'Compra', Factura: d.Factura || '', Proveedor: d.Proveedor || '',
    Codigo: cod, Tipo: prod.Tipo, Cantidad: hasta - desde + 1, FolioDesde: desde, FolioHasta: hasta,
    Neto: neto, IVA: iva, Total: neto + iva, Usuario: d.Usuario || '', Registrado: new Date(),
    Estado: 'Vigente', MotivoAnulacion: ''
  });
  auditar(d.Usuario, 'Ingreso', (d.Origen || 'Compra') + ' ' + cod + ' folios ' + desde + '–' + hasta +
    (d.Factura ? ', factura ' + d.Factura : ''));
  return { ok: true };
}

function anular(nombre, d) {
  checkPin(d);
  if (!d.motivo) throw new Error('Indica el motivo de la anulación');
  const filas = leer(nombre);
  const r = filas.find(x => String(x.ID) === String(d.ID));
  if (!r) throw new Error('Registro no encontrado');
  if (!vigente(r)) throw new Error('El registro ya estaba anulado');
  if (nombre === 'Compras') {
    const dep = leer('Entregas').find(e => vigente(e) && String(e.Codigo) === String(r.Codigo) &&
      seSolapan(Number(r.FolioDesde), Number(r.FolioHasta), Number(e.FolioDesde), Number(e.FolioHasta)));
    if (dep) throw new Error('No se puede anular: hay folios de este ingreso entregados el ' + dep.Fecha +
      ' a ' + dep.Responsable + '. Anula primero esa entrega.');
  }
  actualizar(nombre, r._row, { Estado: 'Anulada', MotivoAnulacion: d.motivo });
  auditar(d.Usuario, 'Anulación ' + (nombre === 'Compras' ? 'ingreso' : 'entrega'),
    r.Codigo + ' folios ' + r.FolioDesde + '–' + r.FolioHasta + '. Motivo: ' + d.motivo);
  return { ok: true };
}

function saveProducto(d) {
  checkPin(d);
  const cod = String(d.Codigo || '').trim();
  if (!cod) throw new Error('Falta el código');
  if (['Tapón', 'Fleje'].indexOf(d.Tipo) < 0) throw new Error('Tipo debe ser Tapón o Fleje');
  const obj = { Codigo: cod, Tipo: d.Tipo, Descripcion: d.Descripcion || '', Prefijo: d.Prefijo || '',
    StockMinimo: Number(d.StockMinimo) || 0, Activo: d.Activo === 'NO' ? 'NO' : 'SI' };
  const r = leer('Productos').find(p => String(p.Codigo) === cod);
  if (r) {
    actualizar('Productos', r._row, { Tipo: obj.Tipo, Descripcion: obj.Descripcion, Prefijo: obj.Prefijo,
      StockMinimo: obj.StockMinimo, Activo: obj.Activo });
  } else {
    agregar('Productos', obj);
  }
  auditar(d.Usuario, r ? 'Editar sello' : 'Nuevo sello', cod + ' (' + obj.Tipo + '), mínimo ' + obj.StockMinimo);
  return { ok: true };
}

function saveConfig(d) {
  checkPin(d);
  if (d.Areas !== undefined) setConfig('Areas', d.Areas);
  if (d.Responsables !== undefined) setConfig('Responsables', d.Responsables);
  if (d.IVA !== undefined) setConfig('IVA', String(d.IVA));
  auditar(d.Usuario, 'Configuración', 'Áreas / responsables / IVA actualizados');
  return { ok: true };
}
