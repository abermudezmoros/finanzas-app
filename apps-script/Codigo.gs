/**
 * Puente entre la app "Registro rápido" y la hoja "Finanzas personales".
 * Se pega en Extensiones > Apps Script de la propia hoja.
 * La clave se guarda en Configuración del proyecto > Propiedades de script, con el nombre CLAVE.
 */

const TIPOS = ['Gasto', 'Ingreso', 'Reembolso', 'Traspaso', 'Inversión'];
const CUENTAS = ['BBVA', 'Trade Republic', 'TR Ahorro', 'Cobee', 'Cartera'];

function claveOk_(clave) {
  const buena = PropertiesService.getScriptProperties().getProperty('CLAVE');
  return buena && clave === buena;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Lectura para la pestaña "Últimos": saldos, resumen del mes y últimos movimientos.
function doGet(e) {
  if (!claveOk_(e.parameter.clave)) return json_({ ok: false, error: 'Clave incorrecta' });
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();

  const saldos = {};
  ss.getSheetByName('Cuentas').getRange('A5:D9').getValues()
    .forEach(r => { if (r[0]) saldos[r[0]] = r[3]; });

  const hoja = ss.getSheetByName('Movimientos');
  const n = hoja.getLastRow() - 1;
  // Cada fila lleva al final su número de fila en la hoja, para poder borrarla desde la app.
  const filas = n > 0 ? hoja.getRange(2, 1, n, 8).getValues()
    .map((r, k) => r.concat(k + 2)).filter(r => r[0] && r[1]) : [];

  const hoy = new Date();
  const mes = Utilities.formatDate(hoy, tz, 'yyyy-MM');
  let ingresos = 0, gastos = 0;
  filas.forEach(r => {
    if (Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM') !== mes) return;
    const imp = Number(r[2]) || 0;
    if (r[1] === 'Ingreso' && r[3] !== 'Saveback') ingresos += imp;
    if (r[1] === 'Gasto') gastos += imp;
    if (r[1] === 'Reembolso') gastos -= imp;
  });

  const ultimos = filas.slice(-20).reverse().map(r => ({
    f: Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM-dd'),
    t: r[1], i: Number(r[2]), cat: r[3], d: r[4], acc: r[5], dest: r[6], met: r[7], fila: r[8]
  }));

  return json_({ ok: true, saldos, mes: { ingresos, gastos, ahorro: ingresos - gastos }, ultimos });
}

// Escritura: recibe uno o varios movimientos (la app los manda en cola si no había cobertura).
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'JSON no válido' }); }
  if (!claveOk_(body.clave)) return json_({ ok: false, error: 'Clave incorrecta' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const cache = CacheService.getScriptCache();
    const ss = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName('Movimientos');

    // Borrado: solo si la fila sigue siendo el mismo movimiento que ve la app.
    if (body.borrar) {
      const b = body.borrar, fila = Number(b.fila);
      if (!(fila >= 2) || fila > hoja.getLastRow()) return json_({ ok: false, error: 'No encontrado' });
      const r = hoja.getRange(fila, 1, 1, 8).getValues()[0];
      const mismo = Utilities.formatDate(new Date(r[0]), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') === b.f &&
        r[1] === b.t && Math.abs(Number(r[2]) - Number(b.i)) < 0.005 && String(r[4]) === String(b.d || '');
      if (!mismo) return json_({ ok: false, error: 'La hoja ha cambiado, actualiza y prueba otra vez' });
      hoja.deleteRow(fila);
      return json_({ ok: true, borrado: fila });
    }
    const guardados = [];

    (body.movimientos || []).forEach(m => {
      // Evita duplicados si la app reintenta un envío que ya llegó.
      if (m.id && cache.get(m.id)) { guardados.push(m.id); return; }
      if (TIPOS.indexOf(m.t) < 0 || CUENTAS.indexOf(m.acc) < 0 || !(Number(m.i) > 0)) return;

      const p = String(m.f).split('-').map(Number);
      const fecha = new Date(p[0], p[1] - 1, p[2]);
      hoja.appendRow([
        fecha, m.t, Number(m.i), m.t === 'Traspaso' ? '' : (m.cat || ''),
        m.d || '', m.acc, m.dest || '', m.met || ''
      ]);
      const fila = hoja.getLastRow();
      hoja.getRange(fila, 1).setNumberFormat('dd/mm/yyyy');
      hoja.getRange(fila, 3).setNumberFormat('#,##0.00 "€"');
      if (m.id) cache.put(m.id, '1', 21600);
      guardados.push(m.id || fila);
    });

    return json_({ ok: true, guardados });
  } finally {
    lock.releaseLock();
  }
}
