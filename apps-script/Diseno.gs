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

function aplicarDiseno() {
  const ss = SpreadsheetApp.getActive();
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
  let corte = ss.getRangeByName('Corte') ? ss.getRangeByName('Corte').getValue() : 25;

  sh.getCharts().forEach(c => sh.removeChart(c));
  sh.clearConditionalFormatRules();
  const todo = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  todo.breakApart(); todo.clearDataValidations(); sh.clear();
  filas_(sh, 50); columnas_(sh, 15);
  base_(sh, C.bg);
  sh.setTabColor(C.navy);
  sh.setFrozenRows(0);

  [24, 200, 125, 125, 125, 125, 125, 120, 30, 190, 120, 120, 120, 170].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  titulo_(sh, 'B', 'Finanzas personales');
  sh.getRange('B2').setValue('El mes empieza el día').setFontColor(C.muted);
  const rc = sh.getRange('C2').setValue(corte).setHorizontalAlignment('center').setFontWeight('bold');
  entrada_(rc);
  ss.setNamedRange('Corte', rc);
  sh.getRange('D2:G2').merge().setFontColor(C.muted)
    .setValue('Ahorro = traspasos a TR Ahorro. Sobrante = lo que pasa al mes siguiente.');

  // Fórmulas por periodo (una fila por mes, crecen solas).
  const M = c => 'Movimientos!$' + c + ':$' + c;
  const W = M('A') + ',">="&desde,' + M('A') + ',"<"&hasta';
  const per = expr => '=LET(origen,DATE(2026,8,Corte),nmes,MAX(1,DATEDIF(origen,TODAY(),"M")+1),' +
    'MAP(SEQUENCE(nmes),LAMBDA(k,LET(desde,EDATE(origen,k-1),hasta,EDATE(origen,k),' + expr + '))))';
  const ING = 'SUMIFS(' + M('C') + ',' + M('B') + ',"Ingreso",' + M('D') + ',"<>Saveback",' + W + ')';
  const GAS = 'SUMIFS(' + M('C') + ',' + M('B') + ',"Gasto",' + W + ')-SUMIFS(' + M('C') + ',' + M('B') + ',"Reembolso",' + W + ')';
  const AHO = 'SUMIFS(' + M('C') + ',' + M('B') + ',"Traspaso",' + M('G') + ',"TR Ahorro",' + W + ')-SUMIFS(' + M('C') + ',' + M('B') + ',"Traspaso",' + M('F') + ',"TR Ahorro",' + W + ')';
  const INV = 'SUMIFS(' + M('C') + ',' + M('B') + ',"Inversión",' + W + ')';
  const LAB = 'IF(Corte=1,TEXT(desde,"mmmm yyyy"),TEXT(desde,"d mmm")&" – "&TEXT(hasta-1,"d mmm yy"))';

  const T0 = 12, T1 = 47; // filas de datos de la tabla mes a mes
  const rg = col => col + T0 + ':' + col + T1;

  // Tarjetas del mes actual (fila 4 etiqueta, 5 valor, 6 tendencia)
  const ultimo = col => '=IFERROR(INDEX(' + rg(col) + ',COUNTA(' + rg('B') + ')),0)';
  sh.getRange('B4').setValue('ESTE MES').setFontColor(C.muted).setFontSize(9).setFontWeight('bold');
  sh.getRange('B5').setFormula(ultimo('B').replace(',0)', ',"")')).setFontSize(13).setFontWeight('bold');
  sh.getRange('B6').setValue('Tendencia de los últimos meses').setFontColor(C.muted).setFontSize(9);
  const tarjetas = [['C', 'INGRESOS', C.tipos['Ingreso'][0]], ['D', 'GASTO REAL', C.tipos['Gasto'][0]],
    ['E', 'AHORRO', C.tipos['Reembolso'][0]], ['F', 'INVERSIÓN', C.tipos['Inversión'][0]], ['G', 'SOBRANTE', C.navy]];
  tarjetas.forEach(([col, et, color]) => {
    sh.getRange(col + '4').setValue(et).setFontColor(C.muted).setFontSize(9).setFontWeight('bold');
    sh.getRange(col + '5').setFormula(ultimo(col)).setNumberFormat(EUR0).setFontSize(18).setFontWeight('bold').setFontColor(color);
    sh.getRange(col + '6').setFormula('=IFERROR(SPARKLINE(FILTER(' + rg(col) + ',' + rg('B') + '<>""),{"charttype","column";"color","' + color + '"}),"")');
  });
  sh.getRange('B4:G6').setBackground(C.card).setHorizontalAlignment('left');
  sh.getRange('B4:G4').setBorder(true, null, null, null, null, null, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(4, 28); sh.setRowHeight(5, 40); sh.setRowHeight(6, 40);

  // Tabla mes a mes
  sh.getRange('B9').setValue('Mes a mes').setFontSize(13).setFontWeight('bold').setFontColor(C.navy);
  sh.getRange('B10:H10').setValues([['Mes', 'Ingresos', 'Gasto real', 'Ahorro', 'Inversión', 'Sobrante', '% ahorro + inv.']]);
  cabecera_(sh.getRange('B10:H10'));
  sh.setRowHeight(10, 34);
  sh.getRange('B11').setValue('Total');
  ['C', 'D', 'E', 'F', 'G'].forEach(col => sh.getRange(col + '11').setFormula('=SUM(' + rg(col) + ')'));
  sh.getRange('H11').setFormula('=IF(C11=0,"",(E11+F11)/C11)');
  sh.getRange('B11:H11').setFontWeight('bold').setBackground('#E9E6E0');
  sh.getRange('B' + T0).setFormula(per(LAB));
  sh.getRange('C' + T0).setFormula(per(ING));
  sh.getRange('D' + T0).setFormula(per(GAS));
  sh.getRange('E' + T0).setFormula(per(AHO));
  sh.getRange('F' + T0).setFormula(per(INV));
  sh.getRange('G' + T0).setFormula('=ARRAYFORMULA(IF(' + rg('B') + '="","",' + rg('C') + '-' + rg('D') + '-' + rg('E') + '-' + rg('F') + '))');
  sh.getRange('H' + T0).setFormula('=ARRAYFORMULA(IF((' + rg('B') + '="")+(' + rg('C') + '=0),"",(' + rg('E') + '+' + rg('F') + ')/' + rg('C') + '))');
  sh.getRange('C11:G' + T1).setNumberFormat(EUR);
  sh.getRange('H11:H' + T1).setNumberFormat('0%');
  sh.getRange('B' + T0 + ':H' + T1).setBackground(C.card)
    .setBorder(null, null, null, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeights(11, T1 - 10, 28);

  // Saldos y cartera (J:K)
  const saldos = [['Saldos', 'Ahora'], ['BBVA', '=Cuentas!D5'], ['Trade Republic', '=Cuentas!D6'], ['TR Ahorro', '=Cuentas!D7'],
    ['Cobee', '=Cuentas!D8'], ['Disponible para gastar', '=Cuentas!D11'], ['Total en cuentas', '=Cuentas!D12']];
  sh.getRange('J4:K10').setValues(saldos);
  cabecera_(sh.getRange('J4:K4'));
  sh.getRange('J5:K10').setBackground(C.card).setBorder(null, null, null, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('K5:K10').setNumberFormat(EUR);
  sh.getRange('J9:K10').setFontWeight('bold');
  sh.getRange('K10').setFontColor(C.navy);

  const cartera = [['Cartera', ''], ['Total aportado', '=Cuentas!D9'], ['Valor de mercado (escríbelo tú)', valor],
    ['Plusvalía', '=K14-K13'], ['Rentabilidad', '=IFERROR(K14/K13-1,0)']];
  sh.getRange('J12:K16').setValues(cartera);
  cabecera_(sh.getRange('J12:K12'));
  sh.getRange('J12:K12').setBackground(C.tipos['Inversión'][0]);
  sh.getRange('J13:K16').setBackground(C.card).setBorder(null, null, null, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('K13:K15').setNumberFormat(EUR);
  sh.getRange('K16').setNumberFormat('0.0%');
  const vc = sh.getRange('K14'); entrada_(vc); ss.setNamedRange('ValorCartera', vc);

  // Gasto por categoría (J:N)
  const C0 = 19;
  sh.getRange('J18:N18').setValues([['Gasto real por categoría', 'Este mes', 'Mes anterior', 'Media', '']]);
  cabecera_(sh.getRange('J18:N18'));
  sh.setRowHeight(18, 34);
  const pre = '=LET(origen,DATE(2026,8,Corte),nmes,MAX(1,DATEDIF(origen,TODAY(),"M")+1),';
  GASTOS.forEach((cat, i) => {
    const r = C0 + i;
    const g = (ini, fin) => 'LET(desde,' + ini + ',hasta,' + fin + ',SUMIFS(' + M('C') + ',' + M('B') + ',"Gasto",' + M('D') + ',$J' + r + ',' + W +
      ')-SUMIFS(' + M('C') + ',' + M('B') + ',"Reembolso",' + M('D') + ',$J' + r + ',' + W + '))';
    sh.getRange('J' + r).setValue(cat);
    sh.getRange('K' + r).setFormula(pre + g('EDATE(origen,nmes-1)', 'EDATE(origen,nmes)') + ')');
    sh.getRange('L' + r).setFormula(pre + g('EDATE(origen,nmes-2)', 'EDATE(origen,nmes-1)') + ')');
    sh.getRange('M' + r).setFormula(pre + 'IF(nmes<2,0,' + g('origen', 'EDATE(origen,nmes-1)') + '/(nmes-1)))');
    sh.getRange('N' + r).setFormula('=SPARKLINE(MAX(K' + r + ',0),{"charttype","bar";"max",MAX($K$' + C0 + ':$M$' + (C0 + GASTOS.length - 1) + ');"color1","' + C.tipos['Gasto'][0] + '"})');
  });
  const CL = C0 + GASTOS.length;
  sh.getRange('J' + CL).setValue('Total');
  ['K', 'L', 'M'].forEach(col => sh.getRange(col + CL).setFormula('=SUM(' + col + C0 + ':' + col + (CL - 1) + ')'));
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
