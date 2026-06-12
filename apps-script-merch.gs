/**
 * ============================================================
 *  SGC MERCH HUB — Backend Apps Script (v1.0)
 * ============================================================
 *  Contiene:
 *   1. SETUP_MIGRATE()  → corre UNA sola vez desde el editor.
 *      Crea la planilla nueva "Merch Hub DB" y migra todo el
 *      histórico desde "Inventario 2025 - 2026".
 *   2. doGet()   → API de lectura (JSONP) para la app.
 *   3. doPost()  → API de escritura (pedidos, importaciones,
 *      transferencias, ajustes, productos).
 *
 *  Deploy: Implementar → Nueva implementación → App web
 *          Ejecutar como: yo / Acceso: cualquier persona
 * ============================================================
 */

var OLD_SS_ID = '1z_mn9bVQGTEf_dIkVHUtLzZ-rw72JNYyr_F77mtSJPE'; // Inventario 2025 - 2026
var SECRET    = 'SGC-merch-2026-x7k';  // debe coincidir con el SECRET del index.html
var CACHE_TTL = 600;                    // 10 min

var LOCS = ['MKT-Q', 'SHOP-Q', 'MKT-N', 'SHOP-N', 'DEPOSITO'];

var MOTIVOS_SEED = [
  'Reunión de padres', 'Admisiones', 'Open Day', 'Graduación / Egresados',
  'Entrevistas de ingreso', 'Visitas escolares', 'Uso personal', 'Regalo a staff',
  'Bienvenida new staff', 'Bienvenida new families', 'Celebraciones internas',
  'Evento deportivo interno', 'Torneo intercolegial', 'Campamento / Viaje',
  'Interhouse', 'Clubs y actividades', 'Regalo institucional',
  'Evento de promoción externa', 'Ferias educativas', 'Charlas / Conferencias',
  'Family Day', 'Kermesse', 'Fiesta de Fin de Año', 'Aniversario institucional',
  'Acción solidaria', 'Reconocimiento especial', 'Premios o sorteos',
  'Steeplechase', 'Founders Day', 'Exchange program', 'Sustainability Summit',
  'International Day', 'Sports Day', 'Asados de Camada', 'Uso interno',
  'The Good School Guide', 'Prensa', 'Venta directa'
];

/* ============================================================
 *  HELPERS GENERALES
 * ============================================================ */

function getDbId_() {
  return PropertiesService.getScriptProperties().getProperty('MERCH_DB_ID');
}

function num_(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  var s = String(v).replace(/[^0-9,.\-]/g, '');
  if (s.indexOf(',') > -1 && s.indexOf('.') > -1) { s = s.replace(/\./g, '').replace(',', '.'); }
  else if (s.indexOf(',') > -1) { s = s.replace(',', '.'); }
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function dstr_(v) {
  // normaliza fecha a 'yyyy-MM-dd HH:mm'
  if (v instanceof Date) return Utilities.formatDate(v, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd HH:mm');
  var s = String(v || '').trim();
  // dd/MM/yyyy [h:mm:ss]
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (m) {
    return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2) +
           ' ' + ('0' + (m[4] || '0')).slice(-2) + ':' + (m[5] || '00');
  }
  return s;
}

function mesKey_(dstr) {
  // 'yyyy-MM-dd...' → 'yyyy-MM'
  var s = String(dstr || '');
  return s.length >= 7 ? s.substring(0, 7) : '';
}

function nowStr_() {
  return Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd HH:mm');
}

function newId_() {
  return Utilities.formatDate(new Date(), 'GMT', 'yyMMddHHmmss') + '-' + Math.floor(Math.random() * 9000 + 1000);
}

function findHeaderRow_(values, mustContain) {
  // busca en las primeras 6 filas una que contenga todos los textos pedidos
  for (var r = 0; r < Math.min(6, values.length); r++) {
    var row = values[r].map(function (c) { return String(c || '').trim(); });
    var ok = mustContain.every(function (txt) {
      return row.some(function (c) { return c.indexOf(txt) > -1; });
    });
    if (ok) return r;
  }
  return -1;
}

function colIdx_(headerRow, txt) {
  for (var i = 0; i < headerRow.length; i++) {
    if (String(headerRow[i] || '').indexOf(txt) > -1) return i;
  }
  return -1;
}

/* ============================================================
 *  1. MIGRACIÓN — correr UNA vez desde el editor
 * ============================================================ */

function SETUP_MIGRATE() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('MERCH_DB_ID')) {
    throw new Error('Ya existe una base migrada (MERCH_DB_ID). Si querés re-migrar, borrá la propiedad MERCH_DB_ID en Configuración del proyecto → Propiedades del script, y borrá la planilla "Merch Hub DB".');
  }

  var old = SpreadsheetApp.openById(OLD_SS_ID);
  var sheets = old.getSheets();
  var log = [];

  // ---- 1. Clasificar hojas del sheet viejo por firma de encabezado ----
  var shInv = null, shStock = null, shPedidos = null, imports = [];
  sheets.forEach(function (sh) {
    var vals = sh.getRange(1, 1, Math.min(6, sh.getLastRow() || 1), Math.max(1, sh.getLastColumn())).getValues();
    if (!shInv && findHeaderRow_(vals, ['Costo unitario']) > -1) { shInv = sh; return; }
    if (!shStock && findHeaderRow_(vals, ['Dif de cantidad']) > -1) { shStock = sh; return; }
    if (!shPedidos && findHeaderRow_(vals, ['Solicitante']) > -1) { shPedidos = sh; return; }
    if (findHeaderRow_(vals, ['Fecha y hora']) > -1) { imports.push(sh); return; }
  });
  if (!shInv) throw new Error('No encontré la hoja Inventario (encabezado "Costo unitario") en el sheet viejo.');
  log.push('Inventario: ' + shInv.getName());
  log.push('Stock: ' + (shStock ? shStock.getName() : 'NO ENCONTRADA'));
  log.push('Pedidos: ' + (shPedidos ? shPedidos.getName() : 'NO ENCONTRADA'));
  log.push('Imports: ' + imports.map(function (s) { return s.getName(); }).join(', '));

  // shop de cada import: por nombre de hoja, si no por orden (1ra=Quilmes, 2da=North)
  var impShops = imports.map(function (sh, i) {
    var n = sh.getName().toLowerCase();
    if (n.indexOf('quil') > -1 || /(^|\W)q(\W|$)/.test(n)) return 'q';
    if (n.indexOf('north') > -1 || n.indexOf('nor') > -1 || /(^|\W)n(\W|$)/.test(n)) return 'n';
    return i === 0 ? 'q' : 'n';
  });
  log.push('Asignación de shops: ' + imports.map(function (s, i) { return s.getName() + ' → ' + (impShops[i] === 'q' ? 'Quilmes' : 'North'); }).join(' | ') + '  ← VERIFICAR');

  // ---- 2. PRODUCTOS desde Inventario ----
  var invVals = shInv.getDataRange().getValues();
  var hr = findHeaderRow_(invVals, ['Costo unitario']);
  var H = invVals[hr];
  var cSku = 0, cNom = colIdx_(H, 'Detalle'), cCosto = colIdx_(H, 'Costo unitario'),
      cCant = colIdx_(H, 'Cantidad'), cInt = colIdx_(H, 'Interno'),
      cPub = colIdx_(H, 'blico'), cP26 = colIdx_(H, '2026');
  // '($) Público' y '($) Público (2026)': cPub agarra el primero que matchee 'blico'
  var productos = [];
  var miscN = 0;
  for (var r = hr + 1; r < invVals.length; r++) {
    var row = invVals[r];
    var nom = String(row[cNom] || '').trim();
    if (!nom) continue;
    var sku = String(row[cSku] || '').trim();
    if (!sku || sku === '-') { miscN++; sku = 'SCGX' + ('0' + miscN).slice(-2); }
    productos.push({
      sku: sku, nombre: nom, costo: num_(row[cCosto]), cantidad: num_(row[cCant]),
      pInt: num_(row[cInt]), pPub: num_(row[cPub]), p26: num_(row[cP26])
    });
  }
  log.push('Productos migrados: ' + productos.length + (miscN ? ' (' + miscN + ' sin SKU → SCGX##)' : ''));
  var prodBySku = {};
  productos.forEach(function (p) { prodBySku[p.sku] = p; });
  var prodByName = {};
  productos.forEach(function (p) { prodByName[p.nombre.toLowerCase()] = p; });

  // ---- 3. MOVIMIENTOS ----
  var movs = []; // [id, fecha, tipo, sku, cantidad, desde, hacia, importe, shop, motivo, usuario, nota]
  var FECHA_MIG = '2025-03-01 00:00';

  function pushMov(fecha, tipo, sku, cant, desde, hacia, importe, shop, motivo, usuario, nota) {
    movs.push([newId_(), fecha, tipo, sku, cant, desde, hacia, importe, shop, motivo, usuario, nota]);
  }

  // 3a. Compras iniciales según distribución de la hoja Stock
  if (shStock) {
    var stVals = shStock.getDataRange().getValues();
    var shr = findHeaderRow_(stVals, ['Dif de cantidad']);
    var SH = stVals[shr];
    var sSku = 0, sCant = colIdx_(SH, 'Cantidad'),
        sMQ = colIdx_(SH, 'MKT - Quilmes'), sSQ = colIdx_(SH, 'Shop Quilmes'),
        sMN = colIdx_(SH, 'MKT - North'), sSN = colIdx_(SH, 'Shop North');
    for (var r2 = shr + 1; r2 < stVals.length; r2++) {
      var sr = stVals[r2];
      var sk = String(sr[sSku] || '').trim();
      if (!sk || !prodBySku[sk]) {
        // fila sin SKU válido: intentar por nombre
        var pn = prodByName[String(sr[1] || '').trim().toLowerCase()];
        if (!pn) continue;
        sk = pn.sku;
      }
      var asign = [['MKT-Q', num_(sr[sMQ])], ['SHOP-Q', num_(sr[sSQ])], ['MKT-N', num_(sr[sMN])], ['SHOP-N', num_(sr[sSN])]];
      var totAsig = 0;
      asign.forEach(function (a) {
        if (a[1] !== 0) { pushMov(FECHA_MIG, 'compra', sk, a[1], '', a[0], 0, '', '', 'migracion', 'Distribución inicial (migración)'); totAsig += a[1]; }
      });
      var dif = num_(sr[sCant]) - totAsig;
      if (dif !== 0) {
        pushMov(FECHA_MIG, 'ajuste', sk, dif, '', 'DEPOSITO', 0, '', '', 'migracion', 'Diferencia detectada en migración — revisar con conteo físico');
      }
    }
    log.push('Movimientos de distribución inicial: ' + movs.length);
  }

  // 3b. Ventas desde las hojas de import (timestamps completos)
  var aliasMap = {}; // CODIGO|shop → sku
  var ventasN = 0, sinSku = 0;
  imports.forEach(function (sh, idx) {
    var shop = impShops[idx];
    var iv = sh.getDataRange().getValues();
    var ihr = findHeaderRow_(iv, ['Fecha y hora']);
    var IH = iv[ihr];
    var iF = colIdx_(IH, 'Fecha'), iC = colIdx_(IH, 'digo'), iD = colIdx_(IH, 'Detalle'),
        iQ = colIdx_(IH, 'Cantidad'), iI = colIdx_(IH, 'Importe'), iS = colIdx_(IH, '#');
    for (var r3 = ihr + 1; r3 < iv.length; r3++) {
      var vr = iv[r3];
      var det = String(vr[iD] || '').trim();
      var fcell = vr[iF];
      var isDate = (fcell instanceof Date) || /^\d{1,2}\/\d{1,2}\/\d{4}/.test(String(fcell || ''));
      if (!isDate || det.indexOf('Total (') === 0 || det.indexOf('Total(') === 0) continue;
      var cant = num_(vr[iQ]); if (!cant) continue;
      var sku2 = String(vr[iS] || '').trim();
      var code = String(vr[iC] || '').trim().toUpperCase();
      if (!sku2 || !prodBySku[sku2]) { sinSku++; continue; }
      if (code) aliasMap[code] = sku2;
      pushMov(dstr_(fcell), 'venta', sku2, cant, shop === 'q' ? 'SHOP-Q' : 'SHOP-N', '', num_(vr[iI]), shop, 'Venta shop', 'migracion', det);
      ventasN++;
    }
  });
  log.push('Ventas migradas: ' + ventasN + (sinSku ? ' (' + sinSku + ' filas sin SKU mapeado, omitidas)' : ''));

  // ---- 4. PEDIDOS + sus movimientos ----
  var pedidos = []; // [id, fecha, solicitante, area, site, motivo, sku, producto, cantidad, estado, lista, valor, usuario, nota]
  if (shPedidos) {
    var pv = shPedidos.getDataRange().getValues();
    var phr = findHeaderRow_(pv, ['Solicitante']);
    var PH = pv[phr];
    var pMes = colIdx_(PH, 'Mes'), pF = pMes - 1,
        pSol = colIdx_(PH, 'Solicitante'), pArea = colIdx_(PH, 'rea'),
        pSite = colIdx_(PH, 'Site'), pMot = colIdx_(PH, 'Motivo'),
        pProd = colIdx_(PH, 'Producto'), pSku = colIdx_(PH, '#'),
        pCant = colIdx_(PH, 'Cantidad'), pEst = colIdx_(PH, 'Estado'),
        pVal = colIdx_(PH, 'Valor'), pVen = colIdx_(PH, 'Ventas');
    if (pF < 0) pF = 0;
    for (var r4 = phr + 1; r4 < pv.length; r4++) {
      var pr = pv[r4];
      var sol = String(pr[pSol] || '').trim();
      var fcell2 = pr[pF];
      var hasDate = (fcell2 instanceof Date) || /^\d{1,2}\/\d{1,2}\/\d{4}/.test(String(fcell2 || ''));
      if (!sol || !hasDate) continue;
      var sku3 = String(pr[pSku] || '').trim();
      var prodNom = String(pr[pProd] || '').trim();
      if ((!sku3 || !prodBySku[sku3]) && prodByName[prodNom.toLowerCase()]) sku3 = prodByName[prodNom.toLowerCase()].sku;
      var cant3 = num_(pr[pCant]);
      var listaRaw = String(pr[pVal] || '').trim();
      var lista = listaRaw.indexOf('Regalo') > -1 ? 'regalo' : (listaRaw.indexOf('Interno') > -1 ? 'interno' : 'publico');
      var site = String(pr[pSite] || '').trim();
      var fped = dstr_(fcell2);
      var pid = newId_();
      pedidos.push([pid, fped, sol, String(pr[pArea] || '').trim(), site,
        String(pr[pMot] || '').trim(), sku3, prodNom, cant3,
        String(pr[pEst] || 'Entregado').trim(), lista, num_(pr[pVen]), 'migracion', '']);
      if (sku3 && cant3) {
        var desde = (site.toLowerCase().indexOf('north') > -1) ? 'MKT-N' : 'MKT-Q';
        pushMov(fped, 'pedido', sku3, cant3, desde, '', num_(pr[pVen]), '',
          String(pr[pMot] || '').trim(), 'migracion', 'Pedido ' + pid + ' — ' + sol);
      }
    }
    log.push('Pedidos migrados: ' + pedidos.length);
  }

  // ---- 5. Crear la planilla nueva ----
  var db = SpreadsheetApp.create('Merch Hub DB');

  var shP = db.getActiveSheet(); shP.setName('PRODUCTOS');
  shP.getRange(1, 1, 1, 9).setValues([['sku', 'nombre', 'costo', 'precio_interno', 'precio_publico', 'precio_2026', 'stock_min', 'activo', 'nota']]);
  if (productos.length) {
    shP.getRange(2, 1, productos.length, 9).setValues(productos.map(function (p) {
      return [p.sku, p.nombre, p.costo, p.pInt, p.pPub, p.p26, 0, 'SI', ''];
    }));
  }

  var shM = db.insertSheet('MOVIMIENTOS');
  shM.getRange(1, 1, 1, 12).setValues([['id', 'fecha', 'tipo', 'sku', 'cantidad', 'desde', 'hacia', 'importe', 'shop', 'motivo', 'usuario', 'nota']]);
  if (movs.length) shM.getRange(2, 1, movs.length, 12).setValues(movs);

  var shPe = db.insertSheet('PEDIDOS');
  shPe.getRange(1, 1, 1, 14).setValues([['id', 'fecha', 'solicitante', 'area', 'site', 'motivo', 'sku', 'producto', 'cantidad', 'estado', 'lista', 'valor', 'usuario', 'nota']]);
  if (pedidos.length) shPe.getRange(2, 1, pedidos.length, 14).setValues(pedidos);

  var shA = db.insertSheet('ALIAS');
  shA.getRange(1, 1, 1, 2).setValues([['codigo_shop', 'sku']]);
  var aliasRows = Object.keys(aliasMap).sort().map(function (k) { return [k, aliasMap[k]]; });
  if (aliasRows.length) shA.getRange(2, 1, aliasRows.length, 2).setValues(aliasRows);
  log.push('Alias de códigos derivados: ' + aliasRows.length);

  var shL = db.insertSheet('LISTAS');
  shL.getRange(1, 1, 1, 2).setValues([['motivo', 'area']]);
  var areas = ['Staff', 'Eventos', 'Admisiones', 'Marketing', 'Dirección', 'Deportes', 'Primaria', 'Secundaria', 'Kinder'];
  var maxL = Math.max(MOTIVOS_SEED.length, areas.length);
  var listRows = [];
  for (var li = 0; li < maxL; li++) listRows.push([MOTIVOS_SEED[li] || '', areas[li] || '']);
  shL.getRange(2, 1, listRows.length, 2).setValues(listRows);

  var shAc = db.insertSheet('Acceso');
  shAc.getRange(1, 1, 2, 1).setValues([['diego.medina@stgeorges.edu.ar'], ['marketing@stgeorges.edu.ar']]);

  var shLog = db.insertSheet('MIGRACION_LOG');
  shLog.getRange(1, 1, log.length + 1, 1).setValues([[nowStr_()]].concat(log.map(function (l) { return [l]; })));

  props.setProperty('MERCH_DB_ID', db.getId());
  Logger.log('============ MIGRACIÓN OK ============');
  log.forEach(function (l) { Logger.log(l); });
  Logger.log('Merch Hub DB: ' + db.getUrl());
  return db.getUrl();
}

/* ============================================================
 *  2. API DE LECTURA — doGet (JSONP)
 * ============================================================ */

function doGet(e) {
  var cb = (e && e.parameter && e.parameter.callback) || 'callback';
  var nocache = e && e.parameter && e.parameter.nocache;
  var cache = CacheService.getScriptCache();

  var json = null;
  if (!nocache) {
    var nChunks = parseInt(cache.get('merch_n') || '0', 10);
    if (nChunks > 0) {
      var parts = [];
      for (var i = 0; i < nChunks; i++) {
        var p = cache.get('merch_' + i);
        if (p === null) { parts = null; break; }
        parts.push(p);
      }
      if (parts) json = parts.join('');
    }
  }

  if (!json) {
    json = JSON.stringify(buildPayload_());
    var chunks = [];
    for (var j = 0; j < json.length; j += 90000) chunks.push(json.substring(j, j + 90000));
    var toCache = { 'merch_n': String(chunks.length) };
    chunks.forEach(function (c, k) { toCache['merch_' + k] = c; });
    try { cache.putAll(toCache, CACHE_TTL); } catch (err) {}
  }

  return ContentService.createTextOutput(cb + '(' + json + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function readTab_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getDataRange().getValues();
  var head = vals[0].map(function (h) { return String(h || '').trim(); });
  return vals.slice(1).map(function (row) {
    var o = {};
    head.forEach(function (h, i) { if (h) o[h] = row[i]; });
    return o;
  });
}

function buildPayload_() {
  var dbId = getDbId_();
  if (!dbId) return { error: 'No hay base migrada. Corré SETUP_MIGRATE() primero.' };
  var ss = SpreadsheetApp.openById(dbId);

  var products = readTab_(ss, 'PRODUCTOS').filter(function (p) { return p.sku; }).map(function (p) {
    return { sku: String(p.sku), nombre: String(p.nombre), costo: num_(p.costo),
             pInt: num_(p.precio_interno), pPub: num_(p.precio_publico), p26: num_(p.precio_2026),
             min: num_(p.stock_min), activo: String(p.activo || 'SI') !== 'NO' };
  });
  var prodBySku = {};
  products.forEach(function (p) { prodBySku[p.sku] = p; });

  var movs = readTab_(ss, 'MOVIMIENTOS');
  var pedidos = readTab_(ss, 'PEDIDOS');
  var aliases = readTab_(ss, 'ALIAS').map(function (a) {
    return { code: String(a.codigo_shop || '').toUpperCase().trim(), sku: String(a.sku || '').trim() };
  }).filter(function (a) { return a.code && a.sku; });

  // ---- stock por ubicación (libro de movimientos) ----
  var stock = {}; // sku → {loc: qty}
  products.forEach(function (p) { stock[p.sku] = { 'MKT-Q': 0, 'SHOP-Q': 0, 'MKT-N': 0, 'SHOP-N': 0, 'DEPOSITO': 0 }; });
  var ventasMes = {}; // 'yyyy-MM' → {q:$, n:$, uq:un, un_:un}
  var ventasSku = {}; // sku → {units, amount, q, n}
  var lastImp = { q: '', n: '' };

  movs.forEach(function (m) {
    var sku = String(m.sku || '').trim();
    if (!stock[sku]) return;
    var qty = num_(m.cantidad);
    var desde = String(m.desde || '').trim(), hacia = String(m.hacia || '').trim();
    if (desde && stock[sku][desde] !== undefined) stock[sku][desde] -= qty;
    if (hacia && stock[sku][hacia] !== undefined) stock[sku][hacia] += qty;
    if (String(m.tipo) === 'venta') {
      var f = dstr_(m.fecha), mk = mesKey_(f), shop = String(m.shop || 'q');
      if (!ventasMes[mk]) ventasMes[mk] = { q: 0, n: 0, uq: 0, un_: 0 };
      var imp = num_(m.importe);
      if (shop === 'n') { ventasMes[mk].n += imp; ventasMes[mk].un_ += qty; }
      else { ventasMes[mk].q += imp; ventasMes[mk].uq += qty; }
      if (!ventasSku[sku]) ventasSku[sku] = { units: 0, amount: 0, q: 0, n: 0 };
      ventasSku[sku].units += qty; ventasSku[sku].amount += imp;
      ventasSku[sku][shop === 'n' ? 'n' : 'q'] += imp;
      if (f > (lastImp[shop === 'n' ? 'n' : 'q'] || '')) lastImp[shop === 'n' ? 'n' : 'q'] = f;
    }
  });

  // ---- pedidos: valor institucional distribuido ----
  var distMotivo = {}, distSku = {};
  var pedidosOut = pedidos.map(function (p) {
    var sku = String(p.sku || '').trim();
    var qty = num_(p.cantidad);
    var prod = prodBySku[sku];
    var pubPrice = prod ? (prod.p26 || prod.pPub) : 0;
    var lista = String(p.lista || 'regalo');
    var valorInst = qty * pubPrice;
    var motivo = String(p.motivo || 'Sin motivo');
    var estado = String(p.estado || '');
    if (estado !== 'Cancelado') {
      if (!distMotivo[motivo]) distMotivo[motivo] = { units: 0, value: 0, charged: 0 };
      distMotivo[motivo].units += qty; distMotivo[motivo].value += valorInst; distMotivo[motivo].charged += num_(p.valor);
      if (sku) {
        if (!distSku[sku]) distSku[sku] = { units: 0, value: 0, charged: 0 };
        distSku[sku].units += qty; distSku[sku].value += valorInst; distSku[sku].charged += num_(p.valor);
      }
    }
    return { id: String(p.id), fecha: dstr_(p.fecha), solicitante: String(p.solicitante || ''),
             area: String(p.area || ''), site: String(p.site || ''), motivo: motivo,
             sku: sku, producto: String(p.producto || (prod ? prod.nombre : '')),
             cantidad: qty, estado: estado, lista: lista, valor: num_(p.valor) };
  });
  pedidosOut.sort(function (a, b) { return a.fecha < b.fecha ? 1 : -1; });

  // ---- compras / inversión por sku ----
  var comprasSku = {};
  movs.forEach(function (m) {
    if (String(m.tipo) === 'compra' || String(m.tipo) === 'ajuste') {
      var sku = String(m.sku || '').trim();
      if (!prodBySku[sku]) return;
      if (!comprasSku[sku]) comprasSku[sku] = 0;
      comprasSku[sku] += num_(m.cantidad) * (String(m.hacia) ? 1 : -1);
    }
  });

  // ---- ROI por producto ----
  var roi = products.map(function (p) {
    var comprado = comprasSku[p.sku] || 0;
    var inversion = comprado * p.costo;
    var v = ventasSku[p.sku] || { units: 0, amount: 0 };
    var d = distSku[p.sku] || { units: 0, value: 0, charged: 0 };
    var retorno = v.amount + d.charged;
    var st = stock[p.sku];
    var rem = st['MKT-Q'] + st['SHOP-Q'] + st['MKT-N'] + st['SHOP-N'] + st['DEPOSITO'];
    return { sku: p.sku, nombre: p.nombre, comprado: comprado, costo: p.costo,
             inversion: inversion, vendidos: v.units, ventas: v.amount,
             regalados: d.units, valorDist: d.value, cobradoPedidos: d.charged,
             stockRem: rem,
             roiCom: inversion > 0 ? (retorno / inversion - 1) : 0,
             roiTotal: inversion > 0 ? ((v.amount + d.value) / inversion - 1) : 0 };
  });

  // ---- stock para tabla ----
  var stockOut = products.map(function (p) {
    var s = stock[p.sku];
    var tot = s['MKT-Q'] + s['SHOP-Q'] + s['MKT-N'] + s['SHOP-N'] + s['DEPOSITO'];
    return { sku: p.sku, nombre: p.nombre, mktQ: s['MKT-Q'], shopQ: s['SHOP-Q'],
             mktN: s['MKT-N'], shopN: s['SHOP-N'], dep: s['DEPOSITO'], total: tot,
             min: p.min, activo: p.activo,
             neg: (s['MKT-Q'] < 0 || s['SHOP-Q'] < 0 || s['MKT-N'] < 0 || s['SHOP-N'] < 0 || s['DEPOSITO'] < 0) };
  });

  var mesesOrd = Object.keys(ventasMes).sort();
  var listas = readTab_(ss, 'LISTAS');
  var motivos = listas.map(function (l) { return String(l.motivo || '').trim(); }).filter(String);
  var areas = listas.map(function (l) { return String(l.area || '').trim(); }).filter(String);

  // Acceso
  var allowed = [];
  try {
    var accSh = ss.getSheetByName('Acceso');
    if (accSh && accSh.getLastRow() > 0) {
      allowed = accSh.getRange('A1:A' + accSh.getLastRow()).getValues()
        .map(function (r) { return String(r[0] || '').toLowerCase().trim(); }).filter(String);
    }
  } catch (err) {}

  return {
    lastUpdate: nowStr_(),
    allowed_emails: allowed,
    products: products,
    stock: stockOut,
    pedidos: pedidosOut,
    aliases: aliases,
    ventas: {
      byMonth: mesesOrd.map(function (k) { return { mes: k, q: ventasMes[k].q, n: ventasMes[k].n, uq: ventasMes[k].uq, un: ventasMes[k].un_ }; }),
      bySku: Object.keys(ventasSku).map(function (k) {
        return { sku: k, nombre: prodBySku[k] ? prodBySku[k].nombre : k, units: ventasSku[k].units, amount: ventasSku[k].amount, q: ventasSku[k].q, n: ventasSku[k].n };
      }).sort(function (a, b) { return b.amount - a.amount; })
    },
    distribuido: {
      byMotivo: Object.keys(distMotivo).map(function (k) {
        return { motivo: k, units: distMotivo[k].units, value: distMotivo[k].value, charged: distMotivo[k].charged };
      }).sort(function (a, b) { return b.value - a.value; })
    },
    roi: roi,
    lists: { motivos: motivos, areas: areas, sites: ['Quilmes', 'North'], locs: LOCS },
    lastImport: lastImp,
    dbUrl: ss.getUrl()
  };
}

/* ============================================================
 *  3. API DE ESCRITURA — doPost
 * ============================================================ */

function doPost(e) {
  var out = { ok: false };
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.secret !== SECRET) { out.error = 'unauthorized'; return jsonOut_(out); }

    var dbId = getDbId_();
    if (!dbId) { out.error = 'sin base migrada'; return jsonOut_(out); }
    var ss = SpreadsheetApp.openById(dbId);

    // validar email contra Acceso
    var email = String(body.email || '').toLowerCase().trim();
    var allowed = [];
    var accSh = ss.getSheetByName('Acceso');
    if (accSh && accSh.getLastRow() > 0) {
      allowed = accSh.getRange('A1:A' + accSh.getLastRow()).getValues()
        .map(function (r) { return String(r[0] || '').toLowerCase().trim(); }).filter(String);
    }
    if (allowed.length && allowed.indexOf(email) === -1) { out.error = 'email no autorizado'; return jsonOut_(out); }

    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var p = body.payload || {};
      switch (body.action) {

        case 'newOrder': {
          var prod = findProduct_(ss, p.sku);
          if (!prod) { out.error = 'producto no encontrado'; break; }
          var qty = num_(p.cantidad);
          if (qty <= 0) { out.error = 'cantidad inválida'; break; }
          var lista = String(p.lista || 'regalo');
          var precio = lista === 'interno' ? num_(prod.precio_interno) : (lista === 'publico' ? (num_(prod.precio_2026) || num_(prod.precio_publico)) : 0);
          var valor = precio * qty;
          var id = newId_();
          var fecha = nowStr_();
          appendRow_(ss, 'PEDIDOS', [id, fecha, String(p.solicitante || ''), String(p.area || ''), String(p.site || 'Quilmes'),
            String(p.motivo || ''), String(p.sku), String(prod.nombre), qty, 'Entregado', lista, valor, email, String(p.nota || '')]);
          var desde = String(p.site || '').toLowerCase().indexOf('north') > -1 ? 'MKT-N' : 'MKT-Q';
          appendRow_(ss, 'MOVIMIENTOS', [newId_(), fecha, 'pedido', String(p.sku), qty, desde, '', valor, '', String(p.motivo || ''), email, 'Pedido ' + id]);
          out.ok = true; out.id = id; out.valor = valor;
          break;
        }

        case 'updateOrderStatus': {
          var shPe = ss.getSheetByName('PEDIDOS');
          var ids = shPe.getRange(2, 1, Math.max(1, shPe.getLastRow() - 1), 1).getValues();
          for (var i = 0; i < ids.length; i++) {
            if (String(ids[i][0]) === String(p.id)) {
              shPe.getRange(i + 2, 10).setValue(String(p.estado));
              out.ok = true; break;
            }
          }
          if (!out.ok) out.error = 'pedido no encontrado';
          break;
        }

        case 'importSales': {
          var shop = p.shop === 'n' ? 'n' : 'q';
          var loc = shop === 'n' ? 'SHOP-N' : 'SHOP-Q';
          var rows = p.rows || [];
          var newAliases = p.newAliases || [];
          newAliases.forEach(function (a) {
            if (a.code && a.sku) appendRow_(ss, 'ALIAS', [String(a.code).toUpperCase().trim(), String(a.sku).trim()]);
          });
          // dedup: clave fecha|sku|cantidad|importe|shop contra las ventas ya registradas
          var existing = {};
          var shMv = ss.getSheetByName('MOVIMIENTOS');
          if (shMv.getLastRow() > 1) {
            shMv.getRange(2, 1, shMv.getLastRow() - 1, 9).getValues().forEach(function (mr) {
              if (String(mr[2]) === 'venta') {
                existing[dstr_(mr[1]) + '|' + mr[3] + '|' + num_(mr[4]) + '|' + num_(mr[7]) + '|' + mr[8]] = true;
              }
            });
          }
          var added = 0, dup = 0;
          rows.forEach(function (r) {
            if (!r.sku || !num_(r.cantidad)) return;
            var key = dstr_(r.fecha) + '|' + String(r.sku) + '|' + num_(r.cantidad) + '|' + num_(r.importe) + '|' + shop;
            if (existing[key]) { dup++; return; }
            existing[key] = true;
            appendRow_(ss, 'MOVIMIENTOS', [newId_(), dstr_(r.fecha), 'venta', String(r.sku), num_(r.cantidad),
              loc, '', num_(r.importe), shop, 'Venta shop', email, String(r.detalle || '')]);
            added++;
          });
          out.ok = true; out.added = added; out.dup = dup;
          break;
        }

        case 'transfer': {
          var qtyT = num_(p.cantidad);
          if (qtyT <= 0 || LOCS.indexOf(p.desde) === -1 || LOCS.indexOf(p.hacia) === -1 || p.desde === p.hacia) { out.error = 'datos inválidos'; break; }
          appendRow_(ss, 'MOVIMIENTOS', [newId_(), nowStr_(), 'transferencia', String(p.sku), qtyT,
            String(p.desde), String(p.hacia), 0, '', '', email, String(p.nota || '')]);
          out.ok = true;
          break;
        }

        case 'adjust': {
          var qtyA = num_(p.cantidad); // puede ser negativa
          if (!qtyA || LOCS.indexOf(p.loc) === -1) { out.error = 'datos inválidos'; break; }
          var row = qtyA > 0
            ? [newId_(), nowStr_(), 'ajuste', String(p.sku), qtyA, '', String(p.loc), 0, '', '', email, String(p.nota || 'Ajuste por conteo')]
            : [newId_(), nowStr_(), 'ajuste', String(p.sku), -qtyA, String(p.loc), '', 0, '', '', email, String(p.nota || 'Ajuste por conteo')];
          appendRow_(ss, 'MOVIMIENTOS', row);
          out.ok = true;
          break;
        }

        case 'addProduct': {
          var shP = ss.getSheetByName('PRODUCTOS');
          var skus = shP.getLastRow() > 1 ? shP.getRange(2, 1, shP.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
          var sku = String(p.sku || '').trim();
          if (!sku) {
            var maxN = 0;
            skus.forEach(function (s) { var m = s.match(/^SCG(\d+)$/); if (m) maxN = Math.max(maxN, parseInt(m[1], 10)); });
            sku = 'SCG' + ('0000' + (maxN + 1)).slice(-4);
          }
          if (skus.indexOf(sku) > -1) { out.error = 'SKU ya existe'; break; }
          appendRow_(ss, 'PRODUCTOS', [sku, String(p.nombre || ''), num_(p.costo), num_(p.pInt), num_(p.pPub), num_(p.p26), num_(p.min), 'SI', String(p.nota || '')]);
          var dist = p.dist || {};
          ['MKT-Q', 'SHOP-Q', 'MKT-N', 'SHOP-N', 'DEPOSITO'].forEach(function (loc2) {
            var q2 = num_(dist[loc2]);
            if (q2 > 0) appendRow_(ss, 'MOVIMIENTOS', [newId_(), nowStr_(), 'compra', sku, q2, '', loc2, 0, '', '', email, 'Alta de producto']);
          });
          out.ok = true; out.sku = sku;
          break;
        }

        case 'updateProduct': {
          var shP2 = ss.getSheetByName('PRODUCTOS');
          var vals = shP2.getRange(2, 1, Math.max(1, shP2.getLastRow() - 1), 1).getValues();
          var found = -1;
          for (var k = 0; k < vals.length; k++) if (String(vals[k][0]) === String(p.sku)) { found = k + 2; break; }
          if (found < 0) { out.error = 'producto no encontrado'; break; }
          var map = { nombre: 2, costo: 3, pInt: 4, pPub: 5, p26: 6, min: 7, activo: 8, nota: 9 };
          Object.keys(map).forEach(function (key) {
            if (p[key] !== undefined) {
              var v = (key === 'nombre' || key === 'activo' || key === 'nota') ? String(p[key]) : num_(p[key]);
              shP2.getRange(found, map[key]).setValue(v);
            }
          });
          out.ok = true;
          break;
        }

        case 'restock': {
          var qtyR = num_(p.cantidad);
          if (qtyR <= 0 || LOCS.indexOf(p.loc) === -1) { out.error = 'datos inválidos'; break; }
          appendRow_(ss, 'MOVIMIENTOS', [newId_(), nowStr_(), 'compra', String(p.sku), qtyR, '', String(p.loc), num_(p.costoTotal), '', '', email, String(p.nota || 'Reposición')]);
          out.ok = true;
          break;
        }

        default:
          out.error = 'acción desconocida: ' + body.action;
      }
    } finally {
      lock.releaseLock();
    }

    // invalidar cache para que el próximo doGet traiga datos frescos
    if (out.ok) {
      var cache = CacheService.getScriptCache();
      var n = parseInt(cache.get('merch_n') || '0', 10);
      var keys = ['merch_n'];
      for (var c = 0; c < n; c++) keys.push('merch_' + c);
      cache.removeAll(keys);
    }
  } catch (err) {
    out.error = String(err);
  }
  return jsonOut_(out);
}

function jsonOut_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function appendRow_(ss, tabName, row) {
  ss.getSheetByName(tabName).appendRow(row);
}

function findProduct_(ss, sku) {
  var rows = readTab_(ss, 'PRODUCTOS');
  for (var i = 0; i < rows.length; i++) if (String(rows[i].sku) === String(sku)) return rows[i];
  return null;
}
