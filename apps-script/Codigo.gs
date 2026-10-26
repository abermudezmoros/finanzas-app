/**
 * Puente entre la app "Registro rápido" y la hoja "Finanzas personales".
 * Se pega en Extensiones > Apps Script de la propia hoja.
 * La clave se guarda en Configuración del proyecto > Propiedades de script, con el nombre CLAVE
 * (Finanzas > Conectar el móvil la crea sola si no existe).
 *
 * @OnlyCurrentDoc  El script solo puede abrir esta hoja, ninguna otra del Drive.
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

// Escritura: recibe uno o varios movimientos (la app los manda en cola si no había cobertura),
// o un movimiento que borrar o editar.
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

    const tz = ss.getSpreadsheetTimeZone();
    // Borrar o editar: solo si la fila sigue siendo el mismo movimiento que ve la app.
    const mismaFila_ = b => {
      const fila = Number(b.fila);
      if (!(fila >= 2) || fila > hoja.getLastRow()) return 0;
      const r = hoja.getRange(fila, 1, 1, 8).getValues()[0];
      const mismo = Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM-dd') === b.f &&
        r[1] === b.t && Math.abs(Number(r[2]) - Number(b.i)) < 0.005 && String(r[4]) === String(b.d || '');
      return mismo ? fila : 0;
    };
    const CAMBIADA = 'La hoja ha cambiado, actualiza y prueba otra vez';
    if (body.borrar) {
      const fila = mismaFila_(body.borrar);
      if (!fila) return json_({ ok: false, error: CAMBIADA });
      hoja.deleteRow(fila);
      return json_({ ok: true, borrado: fila });
    }
    const nombres = leeCuentas_(ss).map(c => c.n);
    const valido_ = m => TIPOS.indexOf(m.t) >= 0 && nombres.indexOf(m.acc) >= 0 && Number(m.i) > 0 &&
      /^\d{4}-\d{2}-\d{2}$/.test(String(m.f));
    // La fecha se crea en la zona horaria de la hoja, para que no cambie de día.
    const filaDe_ = m => [
      Utilities.parseDate(String(m.f), tz, 'yyyy-MM-dd'), m.t, Number(m.i), m.t === 'Traspaso' ? '' : (m.cat || ''),
      m.d || '', m.acc, m.dest || '', m.met || ''
    ];
    const formatos_ = fila => {
      hoja.getRange(fila, 1).setNumberFormat('dd/mm/yyyy');
      hoja.getRange(fila, 3).setNumberFormat('#,##0.00 "€"');
    };
    if (body.editar) {
      const fila = mismaFila_(body.editar);
      if (!fila) return json_({ ok: false, error: CAMBIADA });
      if (!valido_(body.mov || {})) return json_({ ok: false, error: 'Datos no válidos' });
      hoja.getRange(fila, 1, 1, 8).setValues([filaDe_(body.mov)]);
      formatos_(fila);
      return json_({ ok: true, editado: fila });
    }
    const guardados = [];

    (body.movimientos || []).forEach(m => {
      // Evita duplicados si la app reintenta un envío que ya llegó.
      if (m.id && cache.get(m.id)) { guardados.push(m.id); return; }
      if (!valido_(m)) return;
      hoja.appendRow(filaDe_(m));
      const fila = hoja.getLastRow();
      formatos_(fila);
      if (m.id) cache.put(m.id, '1', 21600);
      guardados.push(m.id || fila);
    });

    return json_({ ok: true, guardados });
  } finally {
    lock.releaseLock();
  }
}

// ---------- Conectar el móvil ----------
// Muestra un QR y un enlace que abren la app con la URL y la clave ya puestas.
const APP_URL = 'https://abermudezmoros.github.io/finanzas-app/';

function claveOCrea_() {
  const props = PropertiesService.getScriptProperties();
  let clave = props.getProperty('CLAVE');
  if (!clave) {
    clave = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
    props.setProperty('CLAVE', clave);
  }
  return clave;
}

// URL de la aplicación web publicada (acaba en /exec): la que pegó el usuario o, si no, la que da Google.
function urlApp_() {
  const pegada = PropertiesService.getScriptProperties().getProperty('URL_APP');
  if (pegada) return pegada;
  let url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  return /^https:\/\/script\.google\.com\/.+\/exec$/.test(url) ? url : '';
}

function conectarMovil() {
  const url = urlApp_();
  const t = HtmlService.createTemplate(DIALOGO_CONECTAR);
  t.url = url;
  t.enlace = url ? APP_URL + '#url=' + encodeURIComponent(url) + '&clave=' + encodeURIComponent(claveOCrea_()) : '';
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(460).setHeight(url ? 600 : 560), 'Conectar el móvil');
}

// La llama el cuadro de diálogo cuando el usuario pega la URL de la aplicación web.
function guardarUrlApp(url) {
  url = String(url || '').trim();
  if (!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec$/.test(url)) {
    throw new Error('La URL tiene que empezar por https://script.google.com/macros/s/ y acabar en /exec');
  }
  PropertiesService.getScriptProperties().setProperty('URL_APP', url);
  return true;
}
function olvidarUrlApp() {
  PropertiesService.getScriptProperties().deleteProperty('URL_APP');
}

// El cuadro de diálogo va aquí dentro para que el script siga siendo solo dos archivos.
const DIALOGO_CONECTAR = `<!doctype html>
<html><head><base target="_top">
<style>
  body{font-family:Arial,sans-serif;color:#1B1D21;font-size:14px;line-height:1.45;margin:0 4px}
  h3{color:#1F3864;margin:4px 0 8px;font-size:15px}
  ol{padding-left:20px;margin:6px 0}li{margin:6px 0}
  #qr{display:flex;justify-content:center;margin:12px 0}
  .aviso{background:#FFF6D6;border:1px solid #E8C766;border-radius:8px;padding:8px 10px;font-size:13px}
  button{background:#1F3864;color:#fff;border:0;border-radius:8px;padding:9px 14px;font-size:14px;cursor:pointer}
  input{width:100%;box-sizing:border-box;padding:8px;font-size:13px;margin:6px 0;border:1px solid #ccc;border-radius:6px}
  .gris{color:#6B6F76;font-size:12px}a{color:#2D6CB5}
</style></head><body>
<? if (url) { ?>
  <h3>1. Escanea este código con la cámara del móvil</h3>
  <div id="qr"></div>
  <h3>2. Se abrirá la app ya conectada</h3>
  <ol>
    <li><b>iPhone:</b> en Safari, botón Compartir (el cuadrado con la flecha) &gt; <b>Añadir a pantalla de inicio</b>.</li>
    <li><b>Android:</b> en Chrome, menú ⋮ &gt; <b>Añadir a pantalla de inicio</b> (o <b>Instalar app</b>).</li>
  </ol>
  <p>¿No puedes escanearlo? Copia el enlace, mándatelo por WhatsApp o correo y ábrelo en el móvil.</p>
  <button onclick="copiar()">Copiar enlace</button> <span id="ok" class="gris"></span>
  <p class="aviso">Este enlace es la llave de tu hoja: quien lo tenga puede ver y apuntar tus movimientos. No se lo pases a nadie.</p>
  <p class="gris"><a href="#" onclick="cambiar();return false">Cambiar la URL de la aplicación web</a></p>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
  <script>
    var enlace = <?!= JSON.stringify(enlace) ?>;
    try { new QRCode(document.getElementById('qr'), { text: enlace, width: 220, height: 220, correctLevel: QRCode.CorrectLevel.L }); }
    catch (e) { document.getElementById('qr').textContent = 'No se pudo dibujar el código: usa el botón Copiar enlace.'; }
    function copiar() {
      var t = document.createElement('textarea'); t.value = enlace; document.body.appendChild(t); t.select();
      document.execCommand('copy'); t.remove(); document.getElementById('ok').textContent = 'Copiado ✓';
    }
    function cambiar() { google.script.run.withSuccessHandler(function(){ google.script.run.conectarMovil(); }).olvidarUrlApp(); }
  </script>
<? } else { ?>
  <h3>Antes hay que publicar la hoja (solo una vez)</h3>
  <ol>
    <li>Menú <b>Extensiones &gt; Apps Script</b>. Se abre otra pestaña.</li>
    <li>Arriba a la derecha, botón azul <b>Implementar &gt; Nueva implementación</b>.</li>
    <li>Junto a «Seleccionar tipo», pulsa la rueda ⚙ y elige <b>Aplicación web</b>.</li>
    <li>En «Quién tiene acceso» elige <b>Cualquier usuario</b>. Lo demás, tal cual.</li>
    <li>Pulsa <b>Implementar</b>. Si pide permisos: <b>Autorizar acceso</b> &gt; tu cuenta &gt; <b>Configuración avanzada</b> &gt; <b>Ir a…</b> &gt; <b>Permitir</b>.</li>
    <li>Debajo de «Aplicación web» verás una URL: pulsa <b>Copiar</b> y pégala aquí:</li>
  </ol>
  <input id="u" placeholder="https://script.google.com/macros/s/…/exec">
  <button onclick="guardar()">Continuar</button> <span id="err" class="gris"></span>
  <script>
    function guardar() {
      document.getElementById('err').textContent = 'Comprobando…';
      google.script.run
        .withSuccessHandler(function(){ google.script.run.conectarMovil(); })
        .withFailureHandler(function(e){ document.getElementById('err').textContent = e.message; })
        .guardarUrlApp(document.getElementById('u').value);
    }
  </script>
<? } ?>
</body></html>`;
