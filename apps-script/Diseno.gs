/**
 * Diseño de la hoja "Finanzas personales".
 * Se ejecuta desde el menú Finanzas > Aplicar diseño (o eligiendo aplicarDiseno y pulsando Ejecutar).
 * Solo cambia el formato y reconstruye el Resumen; no toca ningún movimiento.
 */

const C = {
  bg: '#F4F2EE', card: '#FFFFFF', text: '#1B1D21', muted: '#6B6F76', line: '#E2DED7',
  navy: '#1F3864', input: '#FFF6D6', inputLine: '#E8C766',
  tipos: {
    'Gasto': ['#C9473F', '#FBE7E5'],
    'Ingreso': ['#2E7D4F', '#E2F1E7'],
    'Reembolso': ['#2D6CB5', '#E1ECF8'],
    'Traspaso': ['#6B6F76', '#EEEEEE'],
    'Inversión': ['#7445B0', '#EEE6F8']
  }
};
const FUENTE = 'Manrope';
const EUR = '#,##0.00 "€"';
const EUR0 = '#,##0 "€"';
const FECHA_HORA = 'dd mmm yyyy"  "hh:mm';
const GASTOS_POR_DEFECTO = ['Restaurantes', 'Supermercado', 'Ocio', 'Deporte', 'Salud', 'Cuidado personal', 'Coche y moto',
  'Viajes', 'Suscripciones', 'Compras y hogar', 'Regalos', 'Vivienda', 'Formación', 'Impuestos', 'Otros'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Finanzas').addItem('Conectar el móvil', 'conectarMovil').addSeparator()
    .addItem('Renombrar cuenta', 'renombrarCuenta').addItem('Aplicar diseño', 'aplicarDiseno')
    .addItem('Crear o reparar las pestañas', 'prepararHoja').addToUi();
}

// ---------- Hoja nueva ----------
// Crea las pestañas que falten con valores de ejemplo (sirve para montar la plantilla desde una hoja vacía)
// y aplica el diseño. En una hoja que ya tiene datos no borra nada: solo añade lo que falte.
function prepararHoja() {
  const ss = SpreadsheetApp.getActive();
  let mov = ss.getSheetByName('Movimientos');
  if (!mov) {
    mov = ss.insertSheet('Movimientos');
    mov.getRange('A1:H1').setValues([['Fecha', 'Tipo', 'Importe', 'Categoría', 'Descripción', 'Cuenta', 'Destino', 'Método']]);
  }
  let cu = ss.getSheetByName('Cuentas');
  if (!cu) {
    cu = ss.insertSheet('Cuentas');
    cu.getRange('A4:F4').setValues([['Cuenta', 'Uso', 'Saldo inicial', 'Saldo actual', 'Nómina abre ciclo', 'Activa']]);
    cu.getRange('A5:F7').setValues([['Banco', 'Día a día', 0, '', 'Sí', 'Sí'], ['Ahorro', 'Ahorro', 0, '', 'No', 'Sí'],
      ['Inversiones', 'Inversión', 0, '', 'No', 'Sí']]);
  }
  if (!ss.getSheetByName('Categorías')) {
    const ca = ss.insertSheet('Categorías');
    const cols = {
      A: ['Tipo'].concat(Object.keys(C.tipos)),
      C: ['Gasto (y Reembolso)'].concat(GASTOS_POR_DEFECTO),
      E: ['Ingreso', 'Nómina', 'Ventas', 'Intereses', 'Otros ingresos'],
      G: ['Método', 'Tarjeta', 'Bizum', 'Transferencia', 'Domiciliado', 'Efectivo'],
      M: ['Inversión', 'Fondos', 'Acciones', 'Otras inversiones']
    };
    Object.keys(cols).forEach(c => ca.getRange(c + '1:' + c + cols[c].length).setValues(cols[c].map(x => [x])));
  }
  if (!ss.getSheetByName('Resumen')) ss.insertSheet('Resumen', 0);
  // La pestaña vacía que trae una hoja nueva («Hoja 1», «Sheet1»…) sobra.
  ss.getSheets().forEach(sh => {
    if (['Movimientos', 'Cuentas', 'Categorías', 'Resumen'].indexOf(sh.getName()) < 0 && sh.getLastRow() === 0) ss.deleteSheet(sh);
  });
  ['Resumen', 'Movimientos', 'Cuentas', 'Categorías'].forEach((n, i) => {
    ss.setActiveSheet(ss.getSheetByName(n)); ss.moveActiveSheet(i + 1);
  });
  aplicarDiseno();
}

// Las hojas en español separan los argumentos con «;» (y las columnas de {…} con «\»).
// Las fórmulas de este archivo están escritas con comas; F_ las adapta si hace falta.
let SEP = ',';
function detectaSeparador_(ss) {
  const sh = ss.insertSheet('_prueba_');
  sh.getRange('A1').setFormula('=SUM(1,2)');
  SpreadsheetApp.flush();
  SEP = sh.getRange('A1').getValue() === 3 ? ',' : ';';
  ss.deleteSheet(sh);
}
function F_(f) {
  if (SEP === ',' || typeof f !== 'string' || f.charAt(0) !== '=') return f;
  let r = '', enTexto = false, llaves = 0;
  for (const c of f) {
    if (c === '"') enTexto = !enTexto;
    if (!enTexto) {
      if (c === '{') llaves++;
      if (c === '}') llaves--;
      if (c === ',') { r += llaves > 0 ? '\\' : ';'; continue; }
    }
    r += c;
  }
  return r;
}

function aplicarDiseno() {
  const ss = SpreadsheetApp.getActive();
  detectaSeparador_(ss);
  disenoMovimientos_(ss.getSheetByName('Movimientos'));
  disenoCuentas_(ss.getSheetByName('Cuentas'));
  disenoCategorias_(ss.getSheetByName('Categorías'));
  construirResumen_(ss, ss.getSheetByName('Resumen'));
  ss.setActiveSheet(ss.getSheetByName('Resumen'));
  ss.toast('Diseño aplicado', 'Finanzas', 4);
}

// ---------- utilidades ----------
function base_(sh, fondo) {
  const r = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  r.setFontFamily(FUENTE).setFontColor(C.text).setFontSize(10).setVerticalAlignment('middle');
  if (fondo) r.setBackground(fondo);
  sh.setHiddenGridlines(true);
}
function cabecera_(r) {
  r.setBackground(C.navy).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);
}
function titulo_(sh, col, texto, sub) {
  sh.getRange(col + '1').setValue(texto).setFontSize(20).setFontWeight('bold').setFontColor(C.navy);
  sh.setRowHeight(1, 52);
  if (sub) sh.getRange(col + '2').setValue(sub).setFontColor(C.muted).setFontSize(10);
}
function entrada_(r) {
  r.setBackground(C.input).setBorder(true, true, true, true, null, null, C.inputLine, SpreadsheetApp.BorderStyle.SOLID);
}
function filas_(sh, n) {
  if (sh.getMaxRows() < n) sh.insertRowsAfter(sh.getMaxRows(), n - sh.getMaxRows());
}
function columnas_(sh, n) {
  if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), n - sh.getMaxColumns());
}

// ---------- Movimientos ----------
function disenoMovimientos_(sh) {
  const N = 3000;
  filas_(sh, N);
  base_(sh, '#FFFFFF');
  sh.setTabColor(C.navy);
  sh.setFrozenRows(1);

  const head = sh.getRange('A1:H1');
  cabecera_(head);
  sh.setRowHeight(1, 38);

  const datos = sh.getRange(2, 1, N - 1, 8);
  sh.setRowHeights(2, N - 1, 28);
  datos.setBorder(null, null, null, null, null, true, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  // Fecha, y la hora solo en las filas que la tienen (lo que se apunta desde la app).
  const fechas = sh.getRange(2, 1, N - 1, 1), tz = ss.getSpreadsheetTimeZone();
  fechas.setNumberFormats(fechas.getValues().map(([v]) =>
    [v instanceof Date && Utilities.formatDate(v, tz, 'HH:mm') !== '00:00' ? FECHA_HORA : 'dd mmm yyyy']));
  sh.getRange(2, 3, N - 1, 1).setNumberFormat(EUR).setHorizontalAlignment('right').setFontWeight('bold');
  sh.getRange(2, 2, N - 1, 1).setFontWeight('bold');
  sh.getRange(2, 6, N - 1, 3).setFontColor(C.muted);

  [140, 105, 110, 150, 360, 130, 130, 115].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // Colores por tipo: Tipo e Importe con su color fuerte, la fila con el suave.
  const reglas = [];
  Object.keys(C.tipos).forEach(t => {
    const [fuerte, suave] = C.tipos[t];
    reglas.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="' + t + '"')
      .setBackground(suave).setFontColor(fuerte).setRanges([sh.getRange(2, 2, N - 1, 2)]).build());
  });
  Object.keys(C.tipos).forEach(t => {
    reglas.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="' + t + '"')
      .setBackground(C.tipos[t][1]).setRanges([sh.getRange(2, 1, N - 1, 8)]).build());
  });
  sh.setConditionalFormatRules(reglas);

  // Cuenta y Destino: desplegable con las cuentas de la hoja Cuentas (una cuenta nueva aparece sola).
  const cuentas = sh.getParent().getSheetByName('Cuentas').getRange('A' + CTA0 + ':A' + CTA1);
  sh.getRange(2, 6, N - 1, 2).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(cuentas, true).setAllowInvalid(false).build());

  // Tipo, Categoría y Método: desplegables con las listas de la hoja Categorías.
  const cat = sh.getParent().getSheetByName('Categorías');
  if (cat) {
    const de = r => SpreadsheetApp.newDataValidation().requireValueInRange(cat.getRange(r), true).setAllowInvalid(true).build();
    sh.getRange(2, 2, N - 1, 1).setDataValidation(de('A2:A6'));
    sh.getRange(2, 4, N - 1, 1).setDataValidation(de('K2:K200'));
    sh.getRange(2, 8, N - 1, 1).setDataValidation(de('G2:G60'));
  }

  if (!sh.getFilter()) sh.getRange(1, 1, N, 8).createFilter();
}

// ---------- Cuentas: la lista de cuentas de la que sale todo lo demás (A5:F24) ----------
// Uso: «Día a día» cuenta para lo disponible y el sobrante, «Ahorro» recibe el ahorro, «Inversión» es la cartera.
// «Nómina abre ciclo» = Sí en la cuenta donde cobras la nómina principal. «Activa» = No la oculta sin perder el historial.
const CTA0 = 5, CTA1 = 24;
const USOS = ['Día a día', 'Ahorro', 'Inversión'];
const CU = c => 'Cuentas!$' + c + '$' + CTA0 + ':$' + c + '$' + CTA1;
// Convierte el «Tipo» de la hoja antigua (Banco, Ahorro, Tarjeta restaurante, Inversión…) en un uso.
function usoDe_(tipo) {
  const t = String(tipo || '');
  if (USOS.indexOf(t) >= 0) return t;
  if (/ahorro/i.test(t)) return 'Ahorro';
  if (/inversi/i.test(t)) return 'Inversión';
  return 'Día a día';
}
function saldoActual_(r) {
  const s = (col, tipo) => 'SUMIFS(Movimientos!$C:$C,Movimientos!$' + col + ':$' + col + ',$A' + r + ',Movimientos!$B:$B,"' + tipo + '")';
  return '=IF($A' + r + '="","",C' + r + '+' + s('F', 'Ingreso') + '+' + s('F', 'Reembolso') + '-' + s('F', 'Gasto') +
    '-' + s('F', 'Traspaso') + '-' + s('F', 'Inversión') + '+' + s('G', 'Traspaso') + '+' + s('G', 'Inversión') + ')';
}
function disenoCuentas_(sh) {
  filas_(sh, CTA1 + 4); columnas_(sh, 9);
  const n = CTA1 - CTA0 + 1;
  const antigua = sh.getRange('B4').getValue() === 'Tipo';
  if (sh.getRange('A11').getValue() === 'Disponible para gastar') sh.getRange('A11:D12').clearContent();
  const sub = sh.getRange('A2').getValue() || 'Saldo inicial: lo que había en cada cuenta el día que empezaste a apuntar.';
  const v = sh.getRange(CTA0, 1, n, 6).getValues();
  const filas = v.map((x, i) => {
    const r = CTA0 + i;
    if (!x[0]) return ['', '', '', saldoActual_(r), '', ''];
    const ciclo = x[4] === 'Sí' || x[4] === 'No' ? x[4] : (antigua && x[1] === 'Banco' ? 'Sí' : 'No');
    return [x[0], usoDe_(x[1]), x[2] === '' ? 0 : x[2], saldoActual_(r), ciclo, x[5] === 'No' ? 'No' : 'Sí'];
  });

  sh.clearConditionalFormatRules();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.getRange('G1:I' + sh.getMaxRows()).clear();
  base_(sh, C.bg);
  sh.setTabColor(C.tipos['Ingreso'][0]);
  titulo_(sh, 'A', 'Cuentas', sub);
  [210, 120, 130, 140, 150, 90, 230, 130].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  sh.getRange('A4:F4').setValues([['Cuenta', 'Uso', 'Saldo inicial', 'Saldo actual', 'Nómina abre ciclo', 'Activa']]);
  cabecera_(sh.getRange('A4:F4'));
  sh.setRowHeight(4, 34);
  sh.setRowHeights(CTA0, n, 30);
  sh.getRange(CTA0, 1, n, 6).clearFormat().setValues(filas.map(f => f.map(F_)));
  const t = sh.getRange(CTA0, 1, n, 6);
  t.setBackground(C.card).setFontWeight('normal').setFontSize(10).setFontColor(C.text)
    .setBorder(null, null, true, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(CTA0, 1, n, 1).setFontWeight('bold');
  sh.getRange(CTA0, 2, n, 1).setFontColor(C.muted);
  sh.getRange(CTA0, 3, n, 2).setNumberFormat(EUR);
  entrada_(sh.getRange(CTA0, 3, n, 1));
  sh.getRange(CTA0, 4, n, 1).setFontWeight('bold').setFontSize(11);
  sh.getRange(CTA0, 5, n, 2).setHorizontalAlignment('center');

  const lista = l => SpreadsheetApp.newDataValidation().requireValueInList(l, true).setAllowInvalid(false).build();
  sh.getRange(CTA0, 2, n, 1).setDataValidation(lista(USOS));
  sh.getRange(CTA0, 5, n, 2).setDataValidation(lista(['Sí', 'No']));

  // Inversión en morado; cuentas inactivas en gris.
  const fila = sh.getRange(CTA0, 1, n, 6);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F' + CTA0 + '="No"')
      .setFontColor('#A5A8AD').setRanges([fila]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B' + CTA0 + '="Inversión"')
      .setBackground(C.tipos['Inversión'][1]).setRanges([fila]).build()
  ]);

  // Totales (los usa el Resumen)
  sh.getRange('G4:H6').setValues([['Totales', ''],
    ['Disponible para gastar', '=SUMIFS(' + CU('D') + ',' + CU('B') + ',"Día a día")'],
    ['Total en cuentas (con ahorro)', '=SUMIFS(' + CU('D') + ',' + CU('B') + ',"Día a día")+SUMIFS(' + CU('D') + ',' + CU('B') + ',"Ahorro")']]
    .map(f => f.map(F_)));
  cabecera_(sh.getRange('G4:H4'));
  sh.getRange('G5:H6').setBackground(C.card).setFontWeight('bold')
    .setBorder(null, null, true, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('H5:H6').setNumberFormat(EUR).setFontSize(12).setFontColor(C.navy);

  sh.getRange('G8:H16').merge().setWrap(true).setVerticalAlignment('top').setFontColor(C.muted).setValue(
    'Para añadir una cuenta, escríbela en la primera fila libre y rellena el resto.\n\n' +
    'Uso: «Día a día» (cuenta para lo disponible y el sobrante), «Ahorro» (los traspasos a ella son ahorro), «Inversión» (la cartera).\n\n' +
    'Nómina abre ciclo: «Sí» en la cuenta donde cobras la nómina principal.\n\n' +
    'Activa: «No» la quita de la app sin perder sus movimientos (por ejemplo, un banco que dejas).\n\n' +
    'Para cambiar el nombre de una cuenta usa Finanzas > Renombrar cuenta: así se cambia también en los movimientos.');
}

// Cambia el nombre de una cuenta en Cuentas y en todos los movimientos (Cuenta y Destino).
function renombrarCuenta() {
  const ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  const sh = ss.getSheetByName('Cuentas');
  const nombres = sh.getRange(CTA0, 1, CTA1 - CTA0 + 1, 1).getValues().map(r => String(r[0])).filter(String);
  const r1 = ui.prompt('Renombrar cuenta', '¿Qué cuenta quieres renombrar? Escribe su nombre exacto:\n' + nombres.join(', '), ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const viejo = r1.getResponseText().trim();
  const i = nombres.indexOf(viejo);
  if (i < 0) { ui.alert('No hay ninguna cuenta llamada «' + viejo + '».'); return; }
  const r2 = ui.prompt('Renombrar cuenta', 'Nuevo nombre para «' + viejo + '»:', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const nuevo = r2.getResponseText().trim();
  if (!nuevo || nombres.indexOf(nuevo) >= 0) { ui.alert('Ese nombre está vacío o ya existe.'); return; }

  const mov = ss.getSheetByName('Movimientos');
  const n = mov.getLastRow() - 1;
  let cambios = 0;
  if (n > 0) {
    const rg = mov.getRange(2, 6, n, 2);
    // Se quita la validación un momento para poder escribir el nombre nuevo antes de que exista en Cuentas.
    const reglas = rg.getDataValidations();
    rg.clearDataValidations();
    const vals = rg.getValues().map(r => r.map(x => (x === viejo ? (cambios++, nuevo) : x)));
    rg.setValues(vals);
    sh.getRange(CTA0 + i, 1).setValue(nuevo);
    rg.setDataValidations(reglas);
  } else sh.getRange(CTA0 + i, 1).setValue(nuevo);
  ui.alert('Hecho: «' + viejo + '» ahora se llama «' + nuevo + '» (' + cambios + ' movimientos cambiados).');
}

// ---------- Categorías ----------
function disenoCategorias_(sh) {
  base_(sh, C.bg);
  sh.setTabColor(C.muted);
  ['A', 'C', 'E', 'G', 'I', 'K', 'M'].forEach(col => {
    cabecera_(sh.getRange(col + '1'));
    sh.setColumnWidth(sh.getRange(col + '1').getColumn(), 170);
    sh.getRange(col + '2:' + col + '40').setBackground(C.card);
  });
  ['B', 'D', 'F', 'H', 'J', 'L'].forEach(col => sh.setColumnWidth(sh.getRange(col + '1').getColumn(), 16));
  sh.setRowHeight(1, 34);
  filas_(sh, 200);
  sh.setRowHeights(2, 39, 26);
  sh.getRange('A10').setBackground(C.bg).setFontColor(C.muted).setFontStyle('italic').setWrap(true);
  sh.getRange('A7:A40').setBackground(C.bg);
  // La columna Cuenta se rellena sola desde la hoja Cuentas (las cuentas se añaden allí).
  sh.getRange('I1').setValue('Cuenta');
  sh.getRange('I2:I24').clearContent();
  sh.getRange('I2').setFormula(F_('=IFERROR(FILTER(' + CU('A') + ',' + CU('A') + '<>""),"")'));
  // «Todas las categorías» se rellena sola con las de Gasto, Ingreso e Inversión (es la lista del desplegable Categoría).
  sh.getRange('K1').setValue('Todas las categorías');
  sh.getRange('K2:K200').clearContent();
  sh.getRange('K2').setFormula(F_('=IFERROR(FILTER({C2:C60;E2:E60;M2:M60},{C2:C60;E2:E60;M2:M60}<>""),"")'));
  sh.getRange('A10').setValue('Para añadir una categoría, escríbela al final de su columna (Gasto, Ingreso o Inversión). ' +
    '«Todas las categorías» y «Cuenta» se rellenan solas.');
}

// ---------- Resumen ----------
function construirResumen_(ss, sh) {
  // Conserva el valor de mercado de la cartera si ya estaba escrito.
  let valor = '';
  const viejo = sh.getRange('D6').getValue() === 'Valor de mercado (actualízalo tú)' ? sh.getRange('E6').getValue() : null;
  const actual = ss.getRangeByName('ValorCartera');
  valor = actual ? actual.getValue() : (viejo !== null ? viejo : '');
  if (ss.getRangeByName('Corte')) ss.removeNamedRange('Corte');

  sh.getCharts().forEach(c => sh.removeChart(c));
  sh.clearConditionalFormatRules();
  const todo = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  todo.breakApart(); todo.clearDataValidations(); sh.clear();
  filas_(sh, 50); columnas_(sh, 15);
  base_(sh, C.bg);
  sh.setTabColor(C.navy);
  sh.setFrozenRows(0);

  [24, 230, 125, 125, 125, 125, 125, 120, 30, 190, 120, 120, 120, 170].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  titulo_(sh, 'B', 'Finanzas personales');
  sh.getRange('B2:G2').merge().setFontColor(C.muted)
    .setValue('Cada ciclo empieza el día que cobras la nómina principal. Ahorro = traspasos a cuentas de ahorro. Del ciclo = ingresos − gastos − ahorro − inversión. Sobrante = lo acumulado que pasa al ciclo siguiente.');

  // Ciclos: empiezan en cada Ingreso «Nómina» en una cuenta con «Nómina abre ciclo» = Sí y duran hasta la siguiente.
  // Solo funciones clásicas (LET/MAP daban error de análisis en esta hoja).
  const M = c => 'Movimientos!$' + c + ':$' + c;
  const W = (d, h) => M('A') + ',">="&' + d + ',' + M('A') + ',"<"&' + h;
  const suma = (tipo, w, extra) => 'SUMIFS(' + M('C') + ',' + M('B') + ',"' + tipo + '",' + (extra ? extra + ',' : '') + w + ')';
  // Suma de un tipo de movimiento cuya Cuenta (F) o Destino (G) es una cuenta con ese uso.
  // (SUMIFS no acepta una lista de cuentas como criterio en Sheets, por eso va fila a fila.)
  const R2 = c => 'Movimientos!$' + c + '$2:$' + c;
  const enUso = (tipo, d, h, col, uso) => 'SUMPRODUCT(ARRAYFORMULA((' + R2('B') + '="' + tipo + '")*(' + R2('A') + '>=' + d +
    ')*(' + R2('A') + '<' + h + ')*ISNUMBER(MATCH(' + R2(col) + ',IF(' + CU('B') + '="' + uso + '",' + CU('A') + '),0))*' + R2('C') + '))';
  const T0 = 12, T1 = 47; // filas de datos de la tabla de ciclos
  const rg = col => col + T0 + ':' + col + T1;
  const SIN_FIN = 'DATE(2100,1,1)';
  sh.getRange('H2').setFormula(F_('=IFERROR(MAX(' + rg('B') + '),"")'))
    .setNumberFormat('"Ciclo actual desde el "d mmm').setFontColor(C.muted);

  const tabla = [];
  for (let r = T0; r <= T1; r++) {
    const d = '$B' + r, h = 'IF($B' + (r + 1) + '="",' + SIN_FIN + ',$B' + (r + 1) + ')', w = W(d, h);
    const si = x => '=IF($B' + r + '="","",' + x + ')';
    tabla.push([
      si(suma('Ingreso', w) + '-' + enUso('Ingreso', d, h, 'F', 'Inversión')),
      si(suma('Gasto', w) + '-' + suma('Reembolso', w)),
      si(enUso('Traspaso', d, h, 'G', 'Ahorro') + '-' + enUso('Traspaso', d, h, 'F', 'Ahorro')),
      si(suma('Inversión', w)),
      si('C' + r + '-D' + r + '-E' + r + '-F' + r),
      // Sobrante acumulado: lo que tenías al empezar en las cuentas del día a día más el resultado de cada ciclo.
      si((r === T0 ? 'SUMIFS(' + CU('C') + ',' + CU('B') + ',"Día a día")' : 'H' + (r - 1)) + '+G' + r)
    ]);
  }

  // Tarjetas del mes actual (fila 4 etiqueta, 5 valor, 6 tendencia)
  const ultimo = col => '=IFERROR(INDEX(' + rg(col) + ',COUNT(' + rg('B') + ')),0)';
  sh.getRange('B4').setValue('ESTE CICLO').setFontColor(C.muted).setFontSize(9).setFontWeight('bold');
  sh.getRange('B5').setFormula(F_(ultimo('B').replace(',0)', ',"")'))).setNumberFormat('mmmm" · desde el "d mmm').setHorizontalAlignment('left').setFontSize(13).setFontWeight('bold');
  sh.getRange('B6').setValue('Tendencia de los últimos ciclos').setFontColor(C.muted).setFontSize(9);
  const tarjetas = [['C', 'INGRESOS', C.tipos['Ingreso'][0]], ['D', 'GASTO REAL', C.tipos['Gasto'][0]],
    ['E', 'AHORRO', C.tipos['Reembolso'][0]], ['F', 'INVERSIÓN', C.tipos['Inversión'][0]], ['G', 'SOBRANTE', C.navy, 'H']];
  tarjetas.forEach(([col, et, color, origen]) => {
    const src = origen || col;
    sh.getRange(col + '4').setValue(et).setFontColor(C.muted).setFontSize(9).setFontWeight('bold');
    sh.getRange(col + '5').setFormula(F_(ultimo(src))).setNumberFormat(EUR0).setFontSize(18).setFontWeight('bold').setFontColor(color);
    sh.getRange(col + '6').setFormula(F_('=IFERROR(SPARKLINE(FILTER(' + rg(src) + ',' + rg('B') + '<>""),{"charttype","column";"color","' + color + '"}),"")'));
  });
  sh.getRange('B4:G6').setBackground(C.card).setHorizontalAlignment('left');
  sh.getRange('B4:G4').setBorder(true, null, null, null, null, null, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(4, 28); sh.setRowHeight(5, 40); sh.setRowHeight(6, 40);

  // Tabla mes a mes
  sh.getRange('B9').setValue('Ciclo a ciclo').setFontSize(13).setFontWeight('bold').setFontColor(C.navy);
  sh.getRange('B10:H10').setValues([['Ciclo (nómina)', 'Ingresos', 'Gasto real', 'Ahorro', 'Inversión', 'Del ciclo', 'Sobrante']]);
  cabecera_(sh.getRange('B10:H10'));
  sh.setRowHeight(10, 34);
  sh.getRange('B11').setValue('Total');
  ['C', 'D', 'E', 'F', 'G'].forEach(col => sh.getRange(col + '11').setFormula(F_('=SUM(' + rg(col) + ')')));
  sh.getRange('H11').setFormula(F_(ultimo('H')));
  sh.getRange('B11:H11').setFontWeight('bold').setBackground('#E9E6E0');
  sh.getRange('B' + T0).setFormula(F_('=IFERROR(SORT(UNIQUE(FILTER(Movimientos!$A$2:$A,Movimientos!$B$2:$B="Ingreso",' +
    'Movimientos!$D$2:$D="Nómina",ISNUMBER(MATCH(Movimientos!$F$2:$F,FILTER(' + CU('A') + ',' + CU('E') + '="Sí"),0))))),"")'));
  sh.getRange('B' + T0 + ':B' + T1).setNumberFormat('mmmm" · desde el "d mmm').setHorizontalAlignment('left');
  sh.getRange('C' + T0 + ':H' + T1).setFormulas(tabla.map(f => f.map(F_)));
  sh.getRange('C11:H' + T1).setNumberFormat(EUR);
  sh.getRange('G' + T0 + ':G' + T1).setFontColor(C.muted);
  sh.getRange('B' + T0 + ':H' + T1).setBackground(C.card)
    .setBorder(null, null, null, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeights(11, T1 - 10, 28);

  // Saldos y cartera (J:K). Las cuentas salen de la hoja Cuentas (sin las de inversión ni las inactivas a cero).
  const S0 = 7, S1 = 16; // huecos para cuentas
  sh.getRange('J4:K6').setValues([['Saldos', 'Ahora'], ['Disponible para gastar', '=Cuentas!$H$5'], ['Total en cuentas', '=Cuentas!$H$6']]
    .map(f => f.map(F_)));
  sh.getRange('J' + S0).setFormula(F_('=IFERROR(FILTER({' + CU('A') + ',' + CU('D') + '},' + CU('A') + '<>"",' +
    CU('B') + '<>"Inversión",(' + CU('F') + '<>"No")+(' + CU('D') + '<>0)),"")'));
  cabecera_(sh.getRange('J4:K4'));
  sh.getRange('J5:K' + S1).setBackground(C.card).setBorder(null, null, null, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('K5:K' + S1).setNumberFormat(EUR);
  sh.getRange('J5:K6').setFontWeight('bold').setBorder(null, null, true, null, null, null, C.navy, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('K5').setFontColor(C.navy);

  const P0 = S1 + 2; // cartera
  const cartera = [['Cartera', ''], ['Total aportado', '=SUMIFS(' + CU('D') + ',' + CU('B') + ',"Inversión")'],
    ['Valor de mercado (escríbelo tú)', valor], ['Plusvalía', '=K' + (P0 + 2) + '-K' + (P0 + 1)],
    ['Rentabilidad', '=IFERROR(K' + (P0 + 2) + '/K' + (P0 + 1) + '-1,0)']];
  sh.getRange('J' + P0 + ':K' + (P0 + 4)).setValues(cartera.map(f => f.map(F_)));
  cabecera_(sh.getRange('J' + P0 + ':K' + P0));
  sh.getRange('J' + P0 + ':K' + P0).setBackground(C.tipos['Inversión'][0]);
  sh.getRange('J' + (P0 + 1) + ':K' + (P0 + 4)).setBackground(C.card).setBorder(null, null, null, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('K' + (P0 + 1) + ':K' + (P0 + 3)).setNumberFormat(EUR);
  sh.getRange('K' + (P0 + 4)).setNumberFormat('0.0%');
  const vc = sh.getRange('K' + (P0 + 2)); entrada_(vc); ss.setNamedRange('ValorCartera', vc);

  // Gasto por categoría (J:N), con las categorías de gasto de la hoja Categorías
  const hc = ss.getSheetByName('Categorías');
  let gastos = hc ? hc.getRange('C2:C60').getValues().map(r => String(r[0]).trim()).filter(String) : [];
  if (!gastos.length) gastos = GASTOS_POR_DEFECTO;
  const H0 = P0 + 6, C0 = H0 + 1;
  filas_(sh, C0 + gastos.length + 3);
  sh.getRange('J' + H0 + ':N' + H0).setValues([['Gasto real por categoría', 'Este ciclo', 'Anterior', 'Media', '']]);
  cabecera_(sh.getRange('J' + H0 + ':N' + H0));
  sh.setRowHeight(H0, 34);
  const A = '$H$2', previo = 'LARGE(' + rg('B') + ',2)', primero = 'MIN(' + rg('B') + ')', cerrados = '(COUNT(' + rg('B') + ')-1)';
  const gastoCat = (r, d, h) => suma('Gasto', W(d, h), M('D') + ',$J' + r) + '-' + suma('Reembolso', W(d, h), M('D') + ',$J' + r);
  const cats = gastos.map((cat, i) => {
    const r = C0 + i;
    return [cat,
      '=IFERROR(' + gastoCat(r, A, SIN_FIN) + ',0)',
      '=IFERROR(' + gastoCat(r, previo, A) + ',0)',
      '=IFERROR((' + gastoCat(r, primero, A) + ')/' + cerrados + ',0)',
      '=SPARKLINE(MAX(K' + r + ',0),{"charttype","bar";"max",MAX($K$' + C0 + ':$M$' + (C0 + gastos.length - 1) + ');"color1","' + C.tipos['Gasto'][0] + '"})'];
  });
  sh.getRange(C0, 10, gastos.length, 5).setValues(cats.map(f => f.map(F_)));
  const CL = C0 + gastos.length;
  sh.getRange('J' + CL).setValue('Total');
  ['K', 'L', 'M'].forEach(col => sh.getRange(col + CL).setFormula(F_('=SUM(' + col + C0 + ':' + col + (CL - 1) + ')')));
  sh.getRange('J' + C0 + ':N' + (CL - 1)).setBackground(C.card)
    .setBorder(null, null, null, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('J' + CL + ':N' + CL).setFontWeight('bold').setBackground('#E9E6E0');
  sh.getRange('K' + C0 + ':M' + CL).setNumberFormat(EUR);
  sh.getRange('J' + C0 + ':J' + CL).setFontWeight('bold');

  // Sobrante acumulado en verde o rojo (solo es rojo si ni con lo arrastrado llegas)
  const rojo = C.tipos['Gasto'][0], verde = C.tipos['Ingreso'][0];
  const rangos = [sh.getRange('G5'), sh.getRange('H11:H' + T1)];
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor(rojo).setRanges(rangos).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor(verde).setRanges(rangos).build()
  ]);
}
