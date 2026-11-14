/* CATÁLOGO VIRTUAL 2026 — Backend v5 (Google Apps Script) */
/* Usuario inicial: admin / admin123  (cámbialo al entrar) */
const SHEET_ID = ''; /* déjalo vacío si este script se creó desde la hoja (Extensiones → Apps Script) */
const VERSION = 'v5';
const H = {
  Productos: ['id','nombre','categoria','precio','unidad','stock','foto','activo','precio_ant','vistas','vendidos'],
  Banners: ['id','titulo','texto','imagen','orden'],
  Usuarios: ['user','hash','rol'],
  Pedidos: ['id','fecha','cliente','tel','items','total','estado'],
  Config: ['clave','valor'],
  Ventas: ['fecha','dia','pedido','cliente','tel','producto','categoria','cantidad','unidad','precio','subtotal','estado']
};
function ss() { return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive(); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'ping') return out({ ok: true, version: VERSION });
  return HtmlService.createHtmlOutputFromFile('index').setTitle('Catálogo')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1');
}
function doPost(e) {
  try { const r = JSON.parse(e.postData.contents); return ContentService.createTextOutput(rpc(r.f, JSON.stringify(r.a || []))).setMimeType(ContentService.MimeType.JSON); }
  catch (err) { return out({ ok: false, error: err.message }); }
}
/* Punto de entrada único: la página llama rpc('funcion', '[args]') y recibe siempre un texto JSON */
function rpc(f, aj) {
  try {
    const FN = { boot, catalogo, config, banners, login, crearPedido, guardarProducto, borrarProducto, ajustarStock, listaProductos, pedidos, estadoPedido, usuarios, guardarUsuario, guardarConfig, guardarBanner, borrarBanner, exportarExcel, visto };
    if (!FN[f]) throw new Error('Función no permitida: ' + f);
    return JSON.stringify({ ok: true, data: FN[f].apply(null, JSON.parse(aj || '[]')) });
  } catch (e) { return JSON.stringify({ ok: false, error: String(e.message || e) }); }
}
function visto(id) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return;
  try {
    const p = rows('Productos').find(x => x.id === id);
    if (p) { p.vistas = Number(p.vistas || 0) + 1; upsert('Productos', 'id', p); }
  } finally { lock.releaseLock(); }
}
function boot() { init(); return { P: catalogo(), CFG: config(), B: banners() }; }
function sh(n) {
  const book = ss();
  let s = book.getSheetByName(n);
  if (!s) { s = book.insertSheet(n); s.appendRow(H[n]); estilo(s, n); }
  return s;
}
function init() {
  Object.keys(H).forEach(sh);
  const ps0 = sh('Productos');
  if (ps0.getLastColumn() < H.Productos.length) { ps0.getRange(1, 1, 1, H.Productos.length).setValues([H.Productos]); estilo(ps0, 'Productos'); }
  const u = sh('Usuarios');
  if (u.getLastRow() < 2) u.appendRow(['admin', hash('admin123'), 'admin']);
}
function hash(p) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, p));
}
function rows(n) {
  const v = sh(n).getDataRange().getValues(); v.shift();
  const tz = Session.getScriptTimeZone();
  return v.map(r => Object.fromEntries(H[n].map((k, i) => [k, r[i] instanceof Date ? Utilities.formatDate(r[i], tz, 'yyyy-MM-dd HH:mm:ss') : r[i]])));
}
function upsert(n, key, o) {
  const s = sh(n), v = s.getDataRange().getValues(), ki = H[n].indexOf(key);
  const row = H[n].map(k => o[k] === undefined ? '' : o[k]);
  for (let i = 1; i < v.length; i++) if (String(v[i][ki]) === String(o[key])) { s.getRange(i + 1, 1, 1, row.length).setValues([row]); return; }
  s.appendRow(row);
}
function auth(t, roles) {
  const r = CacheService.getScriptCache().get('t_' + t);
  if (!r || roles.indexOf(r) < 0) throw new Error('Sin permiso para esta acción');
  return r;
}
/* ---- público ---- */
function catalogo() {
  return rows('Productos').filter(p => p.activo !== false && String(p.activo).toUpperCase() !== 'FALSE');
}
function config() {
  const c = {}; rows('Config').forEach(r => c[r.clave] = r.valor); return c;
}
function login(user, pass) {
  const u = rows('Usuarios').find(x => x.user === user && x.hash === hash(pass));
  if (!u) throw new Error('Usuario o clave incorrectos');
  const t = Utilities.getUuid();
  CacheService.getScriptCache().put('t_' + t, u.rol, 21600);
  return { token: t, rol: u.rol, user: u.user };
}
function crearPedido(o) {
  const ps = rows('Productos'), id = 'P' + Date.now();
  o.items.forEach(i => {
    const p = ps.find(x => x.id === i.id);
    if (p) { p.stock = Math.max(0, Number(p.stock) - i.qty); p.vendidos = Number(p.vendidos || 0) + i.qty; upsert('Productos', 'id', p); }
  });
  upsert('Pedidos', 'id', { id, fecha: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'), cliente: o.cliente, tel: o.tel, items: JSON.stringify(o.items), total: o.total, estado: 'Nuevo' });
  const f = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const vs = sh('Ventas'), filas = o.items.map(i => {
    const p = ps.find(x => x.id === i.id) || {};
    return [f, f.slice(0, 10), id, o.cliente, "'" + o.tel, i.nombre, p.categoria || '', i.qty, i.unidad, i.precio, i.qty * i.precio, 'Nuevo'];
  });
  if (filas.length) vs.getRange(vs.getLastRow() + 1, 1, filas.length, 12).setValues(filas);
  return id;
}
/* ---- con rol ---- */
function guardarProducto(t, p) {
  auth(t, ['admin']);
  if (!p.id) p.id = 'p' + Date.now();
  const ex = rows('Productos').find(x => x.id === p.id);
  if (ex) { p.vistas = ex.vistas; p.vendidos = ex.vendidos; }
  if (p.img) {
    const m = p.img.match(/^data:(.+);base64,(.+)$/);
    const f = DriveApp.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], p.id + '.jpg'));
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    p.foto = 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w800';
  }
  delete p.img;
  upsert('Productos', 'id', p);
  return p.id;
}
function borrarProducto(t, id) {
  auth(t, ['admin']);
  const p = rows('Productos').find(x => x.id === id);
  if (p) { p.activo = false; upsert('Productos', 'id', p); }
}
function ajustarStock(t, id, stock) {
  auth(t, ['admin', 'bodega']);
  const p = rows('Productos').find(x => x.id === id);
  p.stock = stock; upsert('Productos', 'id', p);
}
function listaProductos(t) { auth(t, ['admin', 'bodega', 'vendedor']); return rows('Productos').filter(p => String(p.activo).toUpperCase() !== 'FALSE'); }
function pedidos(t) { auth(t, ['admin', 'vendedor']); return rows('Pedidos').reverse(); }
function estadoPedido(t, id, estado) {
  auth(t, ['admin', 'vendedor']);
  const o = rows('Pedidos').find(x => x.id === id); o.estado = estado; upsert('Pedidos', 'id', o); syncVentas(id, estado);
}
function usuarios(t) { auth(t, ['admin']); return rows('Usuarios').map(u => ({ user: u.user, rol: u.rol })); }
function guardarUsuario(t, user, pass, rol) {
  auth(t, ['admin']);
  const ex = rows('Usuarios').find(x => x.user === user);
  upsert('Usuarios', 'user', { user, hash: pass ? hash(pass) : ex.hash, rol });
}
function subir(data, name) {
  const m = data.match(/^data:(.+);base64,(.+)$/);
  const f = DriveApp.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name));
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1200';
}
function guardarConfig(t, c) {
  auth(t, ['admin']);
  Object.keys(c).forEach(k => {
    let v = c[k];
    if (String(v).indexOf('data:') === 0) v = subir(v, k);
    upsert('Config', 'clave', { clave: k, valor: v });
  });
}
/* ---- banners del carrusel ---- */
function banners() {
  return rows('Banners').sort((a, b) => Number(a.orden) - Number(b.orden));
}
function guardarBanner(t, b) {
  auth(t, ['admin']);
  if (!b.id) b.id = 'b' + Date.now();
  if (b.img) { b.imagen = subir(b.img, b.id); delete b.img; }
  upsert('Banners', 'id', b);
}
/* ---- formato de base de datos + Dashboard ---- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Catálogo')
    .addItem('Dar formato y crear Dashboard', 'armarHojas')
    .addItem('Reconstruir ventas desde pedidos', 'reconstruirVentas')
    .addToUi();
}
const W = {
  Productos: [90, 210, 130, 80, 120, 70, 300, 70, 110, 80, 80],
  Banners: [90, 200, 320, 320, 70],
  Usuarios: [140, 320, 110],
  Pedidos: [120, 150, 170, 130, 420, 90, 120],
  Config: [150, 380],
  Ventas: [150, 95, 120, 170, 120, 210, 130, 80, 100, 80, 90, 120]
};
const FMT = {
  Productos: { D: '#,##0.00', I: '#,##0.00' },
  Pedidos: { B: 'yyyy-mm-dd hh:mm', E: '@', F: '#,##0.00' },
  Ventas: { A: 'yyyy-mm-dd hh:mm', B: 'yyyy-mm-dd', J: '#,##0.00', K: '#,##0.00' }
};
function estilo(s, n) {
  const cols = H[n].length, mr = Math.max(s.getMaxRows(), 2);
  s.setHiddenGridlines(true);
  s.setFrozenRows(1);
  s.setRowHeight(1, 34);
  s.setTabColor(n === 'Ventas' ? '#1B1A1F' : '#E8795A');
  (W[n] || []).forEach((w, i) => s.setColumnWidth(i + 1, w));
  const all = s.getRange(1, 1, mr, cols);
  all.setVerticalAlignment('middle').setFontSize(10)
    .setBorder(true, true, true, true, true, true, '#E2E8E1', SpreadsheetApp.BorderStyle.SOLID);
  if (!s.getBandings().length) {
    all.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
      .setHeaderRowColor('#1B1A1F').setFirstRowColor('#FFFFFF').setSecondRowColor('#FDEEE9');
  }
  s.getRange(1, 1, 1, cols).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  const f = FMT[n] || {};
  Object.keys(f).forEach(l => s.getRange(2, l.charCodeAt(0) - 64, mr - 1, 1).setNumberFormat(f[l]));
  if (!s.getFilter()) all.createFilter();
  const dv = (col, list) => s.getRange(2, col, mr - 1, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(list, true).build());
  const est = ['Nuevo', 'En preparación', 'Entregado', 'Cancelado'];
  if (n === 'Pedidos') dv(7, est);
  if (n === 'Ventas') dv(12, est);
  if (n === 'Usuarios') dv(3, ['admin', 'vendedor', 'bodega']);
  const rl = [];
  if (n === 'Productos') {
    rl.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER($F2),$F2<=5)')
      .setFontColor('#B3382C').setBold(true).setRanges([s.getRange(2, 6, mr - 1, 1)]).build());
  }
  if (n === 'Pedidos' || n === 'Ventas') {
    const rg = [s.getRange(2, n === 'Pedidos' ? 7 : 12, mr - 1, 1)];
    [['Nuevo', '#DCEBFF', '#1D4F91'], ['En preparación', '#FFF1CC', '#8A5A00'], ['Entregado', '#DDF3E4', '#1E6B3A'], ['Cancelado', '#FBE4E1', '#B3382C']]
      .forEach(([t, bg, fg]) => rl.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t).setBackground(bg).setFontColor(fg).setRanges(rg).build()));
  }
  if (rl.length) s.setConditionalFormatRules(rl);
}
function syncVentas(id, estado) {
  const vs = sh('Ventas'), v = vs.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][2]) === String(id)) vs.getRange(i + 1, 12).setValue(estado);
}
function reconstruirVentas() {
  const vs = sh('Ventas'), ps = rows('Productos'), out = [];
  if (vs.getLastRow() > 1) vs.getRange(2, 1, vs.getLastRow() - 1, 12).clearContent();
  rows('Pedidos').forEach(o => {
    const f = String(o.fecha);
    JSON.parse(o.items || '[]').forEach(i => {
      const p = ps.find(x => x.id === i.id) || {};
      out.push([f, f.slice(0, 10), o.id, o.cliente, "'" + o.tel, i.nombre, p.categoria || '', i.qty, i.unidad, i.precio, i.qty * i.precio, o.estado]);
    });
  });
  if (out.length) vs.getRange(2, 1, out.length, 12).setValues(out);
}
function armarHojas() {
  Object.keys(H).forEach(n => estilo(sh(n), n));
  reconstruirVentas();
  dashboard();
  ss().toast('Hojas con formato y Dashboard listos', 'Catálogo', 5);
}
function dashboard() {
  const sp = ss();
  const c = sp.getSheetByName('Calculos') || sp.insertSheet('Calculos');
  c.clear();
  const K = 'Ventas!$K$2:$K', B = 'Ventas!$B$2:$B', L = 'Ventas!$L$2:$L', U = 'Ventas!$H$2:$H', P = 'Ventas!$F$2:$F';
  const heads = { A1: 'Día', B1: 'Ventas', D1: 'Producto', E1: 'Ventas', G1: 'Categoría', H1: 'Ventas', J1: 'Estado', K1: 'Pedidos', M1: 'Producto', N1: 'Stock', O1: 'Vendido', Q1: 'Período', R1: 'Ventas' };
  Object.keys(heads).forEach(a => c.getRange(a).setValue(heads[a]).setFontWeight('bold'));
  c.getRange('A2').setFormula(`=IFERROR(SORT(QUERY(Ventas!A2:L,"select B, sum(K) where L<>'Cancelado' and B is not null group by B order by B desc limit 30",0),1,TRUE),"")`);
  c.getRange('A2:A31').setNumberFormat('dd/mm');
  c.getRange('D2').setFormula(`=IFERROR(QUERY(Ventas!A2:L,"select F, sum(K) where L<>'Cancelado' and F is not null group by F order by sum(K) desc limit 10",0),"")`);
  c.getRange('G2').setFormula(`=IFERROR(QUERY(Ventas!A2:L,"select G, sum(K) where L<>'Cancelado' and G is not null group by G order by sum(K) desc limit 10",0),"")`);
  c.getRange('J2').setFormula(`=IFERROR(QUERY(Pedidos!A2:G,"select G, count(A) where A is not null and G is not null group by G",0),"")`);
  c.getRange('M2').setFormula(`=IFERROR(QUERY(Productos!A2:I,"select B, F where B is not null and H=true limit 15",0),"")`);
  for (let r = 2; r <= 16; r++) c.getRange('O' + r).setFormula(`=IF(M${r}="","",SUMIFS(${U},${P},M${r},${L},"<>Cancelado"))`);
  c.getRange('Q2:Q5').setValues([['Mes anterior'], ['Este mes'], ['7 días previos'], ['Últimos 7 días']]);
  c.getRange('R2').setFormula(`=SUMIFS(${K},${B},">="&DATE(YEAR(TODAY()),MONTH(TODAY())-1,1),${B},"<"&DATE(YEAR(TODAY()),MONTH(TODAY()),1),${L},"<>Cancelado")`);
  c.getRange('R3').setFormula(`=SUMIFS(${K},${B},">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),${L},"<>Cancelado")`);
  c.getRange('R4').setFormula(`=SUMIFS(${K},${B},">="&(TODAY()-13),${B},"<"&(TODAY()-6),${L},"<>Cancelado")`);
  c.getRange('R5').setFormula(`=SUMIFS(${K},${B},">="&(TODAY()-6),${L},"<>Cancelado")`);

  c.getRange('T1').setValue('Producto').setFontWeight('bold'); c.getRange('U1').setValue('Vistas').setFontWeight('bold');
  c.getRange('W1').setValue('Producto').setFontWeight('bold'); c.getRange('X1').setValue('Comprados').setFontWeight('bold');
  c.getRange('T2').setFormula(`=IFERROR(QUERY(Productos!A2:K,"select B, J where B is not null and J > 0 order by J desc limit 10",0),"")`);
  c.getRange('W2').setFormula(`=IFERROR(QUERY(Productos!A2:K,"select B, K where B is not null and K > 0 order by K desc limit 10",0),"")`);
  const d = sp.getSheetByName('Dashboard') || sp.insertSheet('Dashboard');
  d.getRange('A1:Z70').breakApart();
  d.clear();
  d.getCharts().forEach(ch => d.removeChart(ch));
  d.setHiddenGridlines(true);
  d.setTabColor('#E8795A');
  d.setColumnWidth(1, 20);
  for (let i = 2; i <= 13; i++) d.setColumnWidth(i, 90);
  d.getRange('B2').setValue('Dashboard de ventas').setFontSize(26).setFontWeight('bold').setFontColor('#1B1A1F');
  d.getRange('B3').setValue('Se actualiza solo con cada pedido. Las ventas canceladas no cuentan.').setFontColor('#66726B');
  d.setRowHeight(6, 44);
  const kp = [
    ['VENTAS TOTALES', `=SUMIFS(${K},${L},"<>Cancelado")`, '#,##0.00'],
    ['VENTAS DE HOY', `=SUMIFS(${K},${B},TODAY(),${L},"<>Cancelado")`, '#,##0.00'],
    ['VENTAS DEL MES', '=Calculos!R3', '#,##0.00'],
    ['PEDIDOS', '=MAX(0,COUNTIFS(Pedidos!A:A,"<>",Pedidos!G:G,"<>Cancelado")-1)', '0'],
    ['TICKET PROMEDIO', '=IFERROR(B6/H6,0)', '#,##0.00'],
    ['STOCK BAJO (≤5)', '=COUNTIFS(Productos!F2:F,"<=5",Productos!B2:B,"<>")', '0']
  ];
  kp.forEach(([lab, f, fmt], i) => {
    const col = 2 + 2 * i;
    d.getRange(5, col, 1, 2).merge().setValue(lab).setFontSize(9).setFontColor('#66726B').setBackground('#FDEEE9').setHorizontalAlignment('center').setVerticalAlignment('bottom');
    d.getRange(6, col, 1, 2).merge().setBackground('#FDEEE9').setFontSize(20).setFontWeight('bold').setFontColor('#1B1A1F').setHorizontalAlignment('center').setVerticalAlignment('middle').setNumberFormat(fmt);
    d.getRange(6, col).setFormula(f);
    d.getRange(5, col, 2, 2).setBorder(true, true, true, true, false, false, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID_THICK);
  });
  const chart = (type, range, title, row, col, h, extra) => {
    const b = d.newChart().setChartType(type).addRange(range).setNumHeaders(1).setPosition(row, col, 0, 0)
      .setOption('title', title).setOption('width', 540).setOption('height', h)
      .setOption('colors', ['#E8795A', '#1B1A1F', '#F2B8A4', '#6B7770', '#B8D4C5'])
      .setOption('titleTextStyle', { fontSize: 14, bold: true });
    Object.keys(extra || {}).forEach(k => b.setOption(k, extra[k]));
    d.insertChart(b.build());
  };
  const T = Charts.ChartType, none = { legend: { position: 'none' } };
  chart(T.COLUMN, c.getRange('A1:B31'), 'Ventas por día (últimos 30 días)', 8, 2, 300, none);
  chart(T.COLUMN, c.getRange('Q1:R5'), 'Comparativo de períodos', 8, 8, 300, none);
  chart(T.BAR, c.getRange('D1:E11'), 'Top 10 productos por ventas', 24, 2, 320, none);
  chart(T.PIE, c.getRange('G1:H11'), 'Ventas por categoría', 24, 8, 320, { pieHole: 0.45, legend: { position: 'right' } });
  chart(T.COLUMN, c.getRange('M1:O16'), 'Stock vs. unidades vendidas', 41, 2, 320, { legend: { position: 'bottom' } });
  chart(T.PIE, c.getRange('J1:K6'), 'Pedidos por estado', 41, 8, 320, { pieHole: 0.45, legend: { position: 'right' } });
  chart(T.BAR, c.getRange('T1:U11'), 'Productos más vistos', 58, 2, 320, none);
  chart(T.BAR, c.getRange('W1:X11'), 'Productos más comprados (unidades)', 58, 8, 320, none);
  c.setTabColor('#6B7770');
  sp.setActiveSheet(d);
  sp.moveActiveSheet(1);
  c.hideSheet();
}
/* ---- ventas: exportar a Excel (.xlsx con varias hojas) ---- */
function exportarExcel(t, desde, hasta) {
  auth(t, ['admin', 'vendedor']);
  const ped = rows('Pedidos').filter(o => {
    const d = String(o.fecha).slice(0, 10);
    return (!desde || d >= desde) && (!hasta || d <= hasta);
  });
  const det = [], dia = {}, prod = {};
  ped.forEach(o => {
    const f = String(o.fecha);
    JSON.parse(o.items || '[]').forEach(i => {
      det.push([f, o.id, o.cliente, "'" + o.tel, o.estado, i.nombre, i.qty, i.unidad, i.precio, i.qty * i.precio]);
      if (o.estado !== 'Cancelado') {
        const d = f.slice(0, 10), x = dia[d] || (dia[d] = { ids: {}, u: 0, t: 0 });
        x.ids[o.id] = 1; x.u += i.qty; x.t += i.qty * i.precio;
        const p = prod[i.nombre] || (prod[i.nombre] = { q: 0, t: 0 });
        p.q += i.qty; p.t += i.qty * i.precio;
      }
    });
  });
  const ss = SpreadsheetApp.create('Reporte ventas');
  const put = (name, head, data) => {
    const s = ss.insertSheet(name);
    s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#E8795A').setFontColor('#ffffff');
    if (data.length) s.getRange(2, 1, data.length, head.length).setValues(data);
    s.setFrozenRows(1); s.autoResizeColumns(1, head.length);
  };
  put('Ventas detalle', ['Fecha', 'Pedido', 'Cliente', 'Teléfono', 'Estado', 'Producto', 'Cantidad', 'Unidad', 'Precio', 'Subtotal'], det);
  put('Ventas por día', ['Fecha', 'Pedidos', 'Unidades', 'Total'], Object.keys(dia).sort().map(d => [d, Object.keys(dia[d].ids).length, dia[d].u, dia[d].t]));
  put('Por producto', ['Producto', 'Cantidad', 'Total'], Object.keys(prod).map(k => [k, prod[k].q, prod[k].t]).sort((a, b) => b[2] - a[2]));
  put('Pedidos', ['ID', 'Fecha', 'Cliente', 'Teléfono', 'Total', 'Estado'], ped.map(o => [o.id, String(o.fecha), o.cliente, "'" + o.tel, o.total, o.estado]));
  put('Productos', H.Productos, rows('Productos').map(p => H.Productos.map(k => p[k])));
  ss.deleteSheet(ss.getSheets()[0]);
  SpreadsheetApp.flush();
  const blob = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
  const out = { name: 'Reporte_ventas_' + (desde || 'inicio') + '_' + (hasta || 'hoy') + '.xlsx', b64: Utilities.base64Encode(blob.getBytes()) };
  DriveApp.getFileById(ss.getId()).setTrashed(true);
  return out;
}
function borrarBanner(t, id) {
  auth(t, ['admin']);
  const s = sh('Banners'), v = s.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === id) { s.deleteRow(i + 1); return; }
}
