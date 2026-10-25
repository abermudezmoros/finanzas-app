/**
 * Puente entre la app "Registro rápido" y la hoja "Finanzas personales".
 * Se pega en Extensiones > Apps Script de la propia hoja.
 * La clave se guarda en Configuración del proyecto > Propiedades de script, con el nombre CLAVE.
 */

const TIPOS = ['Gasto', 'Ingreso', 'Reembolso', 'Traspaso', 'Inversión'];

// Las cuentas salen de la hoja Cuentas (A5:F24): nombre, uso, saldo inicial, saldo actual, nómina abre ciclo, activa.
// También entiende la hoja antigua (A5:D9 con «Tipo» en la columna B) hasta que se aplique el diseño nuevo.
function leeCuentas_(ss) {
  const sh = ss.getSheetByName('Cuentas');
  const antigua = sh.getRange('B4').getValue() === 'Tipo';
  return sh.getRange(antigua ? 'A5:F9' : 'A5:F24').getValues().filter(r => r[0]).map(r => {
    const b = String(r[1] || '');
    return {
      n: String(r[0]),
      uso: /ahorro/i.test(b) ? 'Ahorro' : /inversi/i.test(b) ? 'Inversión' : 'Día a día',
      saldo: Number(r[3]) || 0,
      ciclo: r[4] ? r[4] === 'Sí' : b === 'Banco',
      activa: r[5] !== 'No'
    };
  });
}

// Categorías y métodos de la hoja Categorías (columnas C, E, G y M).
function leeListas_(ss) {
  const sh = ss.getSheetByName('Categorías');
  if (!sh) return null;
  const col = c => sh.getRange(c + '2:' + c + '60').getValues().map(r => String(r[0]).trim()).filter(String);
  return { Gasto: col('C'), Ingreso: col('E'), 'Inversión': col('M'), metodos: col('G') };
}

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

  const cuentas = leeCuentas_(ss);
  const uso = {};
  cuentas.forEach(c => { uso[c.n] = c.uso; });
  // Saldos: las cuentas activas y las inactivas que aún tienen dinero.
  const saldos = {};
  cuentas.forEach(c => { if (c.activa || Math.abs(c.saldo) >= 0.005) saldos[c.n] = c.saldo; });
  const abreCiclo = cuentas.filter(c => c.ciclo).map(c => c.n);

  const hoja = ss.getSheetByName('Movimientos');
  const n = hoja.getLastRow() - 1;
  // Cada fila lleva al final su número de fila en la hoja, para poder borrarla desde la app.
  const filas = n > 0 ? hoja.getRange(2, 1, n, 8).getValues()
    .map((r, k) => r.concat(k + 2)).filter(r => r[0] && r[1]) : [];

  // Ciclo actual: desde la última nómina cobrada en una cuenta con «Nómina abre ciclo» = Sí.
  let inicio = null;
  filas.forEach(r => {
    if (r[1] === 'Ingreso' && r[3] === 'Nómina' && abreCiclo.indexOf(r[5]) >= 0) {
      const d = new Date(r[0]);
      if (!inicio || d > inicio) inicio = d;
    }
  });
  let ingresos = 0, gastos = 0, ahorro = 0, inversion = 0;
  filas.forEach(r => {
    if (inicio && new Date(r[0]) < inicio) return;
    const imp = Number(r[2]) || 0;
    if (r[1] === 'Ingreso' && uso[r[5]] !== 'Inversión') ingresos += imp;
    if (r[1] === 'Gasto') gastos += imp;
    if (r[1] === 'Reembolso') gastos -= imp;
    if (r[1] === 'Traspaso' && uso[r[6]] === 'Ahorro') ahorro += imp;
    if (r[1] === 'Traspaso' && uso[r[5]] === 'Ahorro') ahorro -= imp;
    if (r[1] === 'Inversión') inversion += imp;
  });
  const desde = inicio ? Utilities.formatDate(inicio, tz, 'yyyy-MM-dd') : '';

  const ultimos = filas.slice(-20).reverse().map(r => ({
    f: Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM-dd'),
    t: r[1], i: Number(r[2]), cat: r[3], d: r[4], acc: r[5], dest: r[6], met: r[7], fila: r[8]
  }));

  // Configuración para la app: cuentas con su uso, categorías y métodos.
  const config = { cuentas: cuentas.map(c => ({ n: c.n, uso: c.uso, ciclo: c.ciclo, activa: c.activa })), listas: leeListas_(ss) };

  const res = { ok: true, saldos, mes: { desde, ingresos, gastos, ahorro, inversion }, ultimos, config };
  // Para la pestaña Análisis: todos los movimientos, en formato compacto
  // [fecha, tipo, importe, categoría, descripción, cuenta, destino, método, fila].
  if (e.parameter.todos) {
    res.todos = filas.map(r => [Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM-dd'),
      r[1], Number(r[2]) || 0, r[3], r[4], r[5], r[6], r[7], r[8]]);
  }
  return json_(res);
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
    const nombres = leeCuentas_(ss).map(c => c.n);

    (body.movimientos || []).forEach(m => {
      // Evita duplicados si la app reintenta un envío que ya llegó.
      if (m.id && cache.get(m.id)) { guardados.push(m.id); return; }
      if (TIPOS.indexOf(m.t) < 0 || nombres.indexOf(m.acc) < 0 || !(Number(m.i) > 0)) return;

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
