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
const GASTOS = ['Restaurantes', 'Supermercado', 'Ocio', 'Deporte', 'Salud', 'Cuidado personal', 'Coche y moto',
  'Viajes', 'Suscripciones', 'Compras y hogar', 'Regalos', 'Vivienda', 'Formación', 'Impuestos', 'Otros'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Finanzas').addItem('Aplicar diseño', 'aplicarDiseno').addToUi();
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
  sh.getRange(2, 1, N - 1, 1).setNumberFormat('dd mmm yyyy');
  sh.getRange(2, 3, N - 1, 1).setNumberFormat(EUR).setHorizontalAlignment('right').setFontWeight('bold');
  sh.getRange(2, 2, N - 1, 1).setFontWeight('bold');
  sh.getRange(2, 6, N - 1, 3).setFontColor(C.muted);

  [110, 105, 110, 150, 360, 130, 130, 115].forEach((w, i) => sh.setColumnWidth(i + 1, w));

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

  if (!sh.getFilter()) sh.getRange(1, 1, N, 8).createFilter();
}

// ---------- Cuentas (las filas 5 a 9 y la columna D las usa la app: no se mueven) ----------
function disenoCuentas_(sh) {
  base_(sh, C.bg);
  sh.setTabColor(C.tipos['Ingreso'][0]);
  titulo_(sh, 'A', 'Cuentas', 'Saldo inicial a 25/08/2026. El saldo actual se calcula solo con los movimientos.');
  sh.setColumnWidth(1, 210); sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 140); sh.setColumnWidth(4, 150);

  cabecera_(sh.getRange('A4:D4'));
  sh.setRowHeight(4, 34);
  sh.setRowHeights(5, 8, 32);
  const t = sh.getRange('A5:D9');
  t.setBackground(C.card).setBorder(null, null, true, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('A5:A9').setFontWeight('bold');
  sh.getRange('B5:B9').setFontColor(C.muted);
  sh.getRange('C5:D9').setNumberFormat(EUR);
  entrada_(sh.getRange('C5:C9'));
  sh.getRange('D5:D9').setFontWeight('bold').setFontSize(11);
  sh.getRange('A9:D9').setBackground(C.tipos['Inversión'][1]);
  sh.getRange('C9').setBackground(C.input);

  const tot = sh.getRange('A11:D12');
  tot.setBackground(C.card).setFontWeight('bold').setBorder(true, null, true, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('B11:B12').setFontWeight('normal').setFontColor(C.muted);
  sh.getRange('D11:D12').setNumberFormat(EUR).setFontSize(12).setFontColor(C.navy);
}

// ---------- Categorías ----------
function disenoCategorias_(sh) {
  base_(sh, C.bg);
  sh.setTabColor(C.muted);
  ['A', 'C', 'E', 'G', 'I', 'K', 'M'].forEach(col => {
    cabecera_(sh.getRange(col + '1'));
    sh.setColumnWidth(sh.getRange(col + '1').getColumn(), 170);
    sh.getRange(col + '2:' + col + '24').setBackground(C.card);
  });
  ['B', 'D', 'F', 'H', 'J', 'L'].forEach(col => sh.setColumnWidth(sh.getRange(col + '1').getColumn(), 16));
  sh.setRowHeight(1, 34);
  sh.setRowHeights(2, 23, 26);
  sh.getRange('A10').setBackground(C.bg).setFontColor(C.muted).setFontStyle('italic').setWrap(true);
  sh.getRange('A7:A24').setBackground(C.bg);
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
    .setValue('Cada ciclo empieza el día que cobras la nómina en BBVA. Ahorro = traspasos a TR Ahorro. Sobrante = lo que pasa al ciclo siguiente.');

  // Ciclos: empiezan en cada Ingreso «Nómina» de BBVA y duran hasta la siguiente (la de Cobee no corta).
  // Solo funciones clásicas (LET/MAP daban error de análisis en esta hoja).
  const M = c => 'Movimientos!$' + c + ':$' + c;
  const W = (d, h) => M('A') + ',">="&' + d + ',' + M('A') + ',"<"&' + h;
  const suma = (tipo, w, extra) => 'SUMIFS(' + M('C') + ',' + M('B') + ',"' + tipo + '",' + (extra ? extra + ',' : '') + w + ')';
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
      si(suma('Ingreso', w, M('D') + ',"<>Saveback"')),
      si(suma('Gasto', w) + '-' + suma('Reembolso', w)),
      si(suma('Traspaso', w, M('G') + ',"TR Ahorro"') + '-' + suma('Traspaso', w, M('F') + ',"TR Ahorro"')),
      si(suma('Inversión', w)),
      si('C' + r + '-D' + r + '-E' + r + '-F' + r),
      '=IF(OR($B' + r + '="",C' + r + '=0),"",(E' + r + '+F' + r + ')/C' + r + ')'
    ]);
  }

  // Tarjetas del mes actual (fila 4 etiqueta, 5 valor, 6 tendencia)
  const ultimo = col => '=IFERROR(INDEX(' + rg(col) + ',COUNT(' + rg('B') + ')),0)';
  sh.getRange('B4').setValue('ESTE CICLO').setFontColor(C.muted).setFontSize(9).setFontWeight('bold');
  sh.getRange('B5').setFormula(F_(ultimo('B').replace(',0)', ',"")'))).setNumberFormat('mmmm" · desde el "d mmm').setHorizontalAlignment('left').setFontSize(13).setFontWeight('bold');
  sh.getRange('B6').setValue('Tendencia de los últimos ciclos').setFontColor(C.muted).setFontSize(9);
  const tarjetas = [['C', 'INGRESOS', C.tipos['Ingreso'][0]], ['D', 'GASTO REAL', C.tipos['Gasto'][0]],
    ['E', 'AHORRO', C.tipos['Reembolso'][0]], ['F', 'INVERSIÓN', C.tipos['Inversión'][0]], ['G', 'SOBRANTE', C.navy]];
  tarjetas.forEach(([col, et, color]) => {
    sh.getRange(col + '4').setValue(et).setFontColor(C.muted).setFontSize(9).setFontWeight('bold');
    sh.getRange(col + '5').setFormula(F_(ultimo(col))).setNumberFormat(EUR0).setFontSize(18).setFontWeight('bold').setFontColor(color);
    sh.getRange(col + '6').setFormula(F_('=IFERROR(SPARKLINE(FILTER(' + rg(col) + ',' + rg('B') + '<>""),{"charttype","column";"color","' + color + '"}),"")'));
  });
  sh.getRange('B4:G6').setBackground(C.card).setHorizontalAlignment('left');
  sh.getRange('B4:G4').setBorder(true, null, null, null, null, null, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(4, 28); sh.setRowHeight(5, 40); sh.setRowHeight(6, 40);

  // Tabla mes a mes
  sh.getRange('B9').setValue('Ciclo a ciclo').setFontSize(13).setFontWeight('bold').setFontColor(C.navy);
  sh.getRange('B10:H10').setValues([['Ciclo (nómina)', 'Ingresos', 'Gasto real', 'Ahorro', 'Inversión', 'Sobrante', '% ahorro + inv.']]);
  cabecera_(sh.getRange('B10:H10'));
  sh.setRowHeight(10, 34);
  sh.getRange('B11').setValue('Total');
  ['C', 'D', 'E', 'F', 'G'].forEach(col => sh.getRange(col + '11').setFormula(F_('=SUM(' + rg(col) + ')')));
  sh.getRange('H11').setFormula(F_('=IF(C11=0,"",(E11+F11)/C11)'));
  sh.getRange('B11:H11').setFontWeight('bold').setBackground('#E9E6E0');
  sh.getRange('B' + T0).setFormula(F_('=IFERROR(SORT(UNIQUE(FILTER(Movimientos!$A$2:$A,Movimientos!$B$2:$B="Ingreso",' +
    'Movimientos!$D$2:$D="Nómina",Movimientos!$F$2:$F="BBVA"))),"")'));
  sh.getRange('B' + T0 + ':B' + T1).setNumberFormat('mmmm" · desde el "d mmm').setHorizontalAlignment('left');
  sh.getRange('C' + T0 + ':H' + T1).setFormulas(tabla.map(f => f.map(F_)));
  sh.getRange('C11:G' + T1).setNumberFormat(EUR);
  sh.getRange('H11:H' + T1).setNumberFormat('0%');
  sh.getRange('B' + T0 + ':H' + T1).setBackground(C.card)
    .setBorder(null, null, null, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeights(11, T1 - 10, 28);

  // Saldos y cartera (J:K)
  const saldos = [['Saldos', 'Ahora'], ['BBVA', '=Cuentas!D5'], ['Trade Republic', '=Cuentas!D6'], ['TR Ahorro', '=Cuentas!D7'],
    ['Cobee', '=Cuentas!D8'], ['Disponible para gastar', '=Cuentas!D11'], ['Total en cuentas', '=Cuentas!D12']];
  sh.getRange('J4:K10').setValues(saldos.map(f => f.map(F_)));
  cabecera_(sh.getRange('J4:K4'));
  sh.getRange('J5:K10').setBackground(C.card).setBorder(null, null, null, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('K5:K10').setNumberFormat(EUR);
  sh.getRange('J9:K10').setFontWeight('bold');
  sh.getRange('K10').setFontColor(C.navy);

  const cartera = [['Cartera', ''], ['Total aportado', '=Cuentas!D9'], ['Valor de mercado (escríbelo tú)', valor],
    ['Plusvalía', '=K14-K13'], ['Rentabilidad', '=IFERROR(K14/K13-1,0)']];
  sh.getRange('J12:K16').setValues(cartera.map(f => f.map(F_)));
  cabecera_(sh.getRange('J12:K12'));
  sh.getRange('J12:K12').setBackground(C.tipos['Inversión'][0]);
  sh.getRange('J13:K16').setBackground(C.card).setBorder(null, null, null, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('K13:K15').setNumberFormat(EUR);
  sh.getRange('K16').setNumberFormat('0.0%');
  const vc = sh.getRange('K14'); entrada_(vc); ss.setNamedRange('ValorCartera', vc);

  // Gasto por categoría (J:N)
  const C0 = 19;
  sh.getRange('J18:N18').setValues([['Gasto real por categoría', 'Este ciclo', 'Anterior', 'Media', '']]);
  cabecera_(sh.getRange('J18:N18'));
  sh.setRowHeight(18, 34);
  const A = '$H$2', previo = 'LARGE(' + rg('B') + ',2)', primero = 'MIN(' + rg('B') + ')', cerrados = '(COUNT(' + rg('B') + ')-1)';
  const gastoCat = (r, d, h) => suma('Gasto', W(d, h), M('D') + ',$J' + r) + '-' + suma('Reembolso', W(d, h), M('D') + ',$J' + r);
  const cats = GASTOS.map((cat, i) => {
    const r = C0 + i;
    return [cat,
      '=IFERROR(' + gastoCat(r, A, SIN_FIN) + ',0)',
      '=IFERROR(' + gastoCat(r, previo, A) + ',0)',
      '=IFERROR((' + gastoCat(r, primero, A) + ')/' + cerrados + ',0)',
      '=SPARKLINE(MAX(K' + r + ',0),{"charttype","bar";"max",MAX($K$' + C0 + ':$M$' + (C0 + GASTOS.length - 1) + ');"color1","' + C.tipos['Gasto'][0] + '"})'];
  });
  sh.getRange(C0, 10, GASTOS.length, 5).setValues(cats.map(f => f.map(F_)));
  const CL = C0 + GASTOS.length;
  sh.getRange('J' + CL).setValue('Total');
  ['K', 'L', 'M'].forEach(col => sh.getRange(col + CL).setFormula(F_('=SUM(' + col + C0 + ':' + col + (CL - 1) + ')')));
  sh.getRange('J' + C0 + ':N' + (CL - 1)).setBackground(C.card)
    .setBorder(null, null, null, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('J' + CL + ':N' + CL).setFontWeight('bold').setBackground('#E9E6E0');
  sh.getRange('K' + C0 + ':M' + CL).setNumberFormat(EUR);
  sh.getRange('J' + C0 + ':J' + CL).setFontWeight('bold');

  // Sobrante en verde o rojo
  const rojo = C.tipos['Gasto'][0], verde = C.tipos['Ingreso'][0];
  const rangos = [sh.getRange('G5'), sh.getRange('G11:G' + T1)];
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor(rojo).setRanges(rangos).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor(verde).setRanges(rangos).build()
  ]);
}
