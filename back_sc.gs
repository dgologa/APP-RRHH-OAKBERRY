function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('OAKBERRY SUPPLY CHAIN')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// --- 1. OBTENCIÓN DE DATOS MAESTROS ---
function obtenerDatosIniciales() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  function leerHoja(nombre, cols) {
    var s = ss.getSheetByName(nombre);
    if (!s || s.getLastRow() < 2) return [];
    return s.getRange(2, 1, s.getLastRow() - 1, cols).getValues();
  }

  // A. Catálogos
  var diario = leerHoja('Catalogo', 4).map(function(r) { 
    return { nombre: String(r[0]), unidad: String(r[1] || "UNIDADES"), proveedor: String(r[2] || "GENERAL"), costo: Number(r[3] || 0) }; 
  }).filter(function(p) { return p.nombre; });

  var general = leerHoja('CatalogoGeneral', 4).map(function(r) { 
    return { nombre: String(r[0]), unidad: String(r[1]), proveedor: String(r[2]), empaques: String(r[3]) }; 
  }).filter(function(p) { return p.nombre; });

  var empaques = leerHoja('Empaques', 3).map(function(r) { return { nombre: String(r[0]), tipo: String(r[1]), valor: Number(r[2]) }; });
  var usuarios = leerHoja('Config', 3).map(function(r) { return { sucursal: String(r[0]), nombre: String(r[1]), id: String(r[2]) }; }).filter(function(u) { return u.sucursal; });
  
  var rawAdm = leerHoja('Admins', 1);
  var admins = [];
  if (rawAdm.length > 0) admins = rawAdm.map(function(r) { return r[0]; }).filter(String);

  var sheetConf = ss.getSheetByName('GlobalConfig');
  var diasPermitidos = [];
  if (sheetConf && sheetConf.getLastRow() >= 2) {
    var dataConf = sheetConf.getRange(2, 1, sheetConf.getLastRow() - 1, 2).getValues();
    var row = dataConf.find(function(r){ return String(r[0]) === 'DIAS_GENERAL'; });
    if (row && row[1] !== "") diasPermitidos = String(row[1]).split(',').map(Number);
  }
  
  var diaJS = new Date().getDay();

  return { diario: diario, general: general, empaques: empaques, usuarios: usuarios, admins: admins, config: { diasGeneral: diasPermitidos, diaActual: diaJS } };
}

// --- 2. LOGINS ---
function loginAdminUser(usuario, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Admins');
  if (!sheet) return { success: false, msg: "Falta hoja Admins" };
  
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  var uInput = String(usuario).trim();
  var pInput = String(password).trim();
  
  var valid = false;
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === uInput && String(data[i][1]).trim() === pInput) { valid = true; break; }
  }
  if (valid) return { success: true };
  return { success: false, msg: "Contraseña incorrecta" };
}

function loginSucursal(sucursal, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Accesos');
  if (!sheet) return { success: false, msg: "Falta hoja Accesos" };
  
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  var valid = false;
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() == String(sucursal).trim() && String(data[i][1]).trim() == String(password).trim()) { valid = true; break; }
  }
  
  if (valid) {
    var sGen = ss.getSheetByName('InventarioGeneral');
    if (!sGen || sGen.getLastRow() < 2) return { success: true, data: [], msg: "Aun no se ha cargado el inventario" };
    var vals = sGen.getRange(2, 1, sGen.getLastRow() - 1, 10).getValues();
    var pendientes = [];
    for (var j = 0; j < vals.length; j++) {
      if (vals[j][1] == sucursal && vals[j][9] == "PENDIENTE") {
        pendientes.push({
          row: j + 2, producto: vals[j][3], total: vals[j][8], colab: vals[j][2],
          fecha: Utilities.formatDate(new Date(vals[j][0]), ss.getSpreadsheetTimeZone(), "dd/MM HH:mm")
        });
      }
    }
    return { success: true, data: pendientes, msg: pendientes.length > 0 ? "" : "Aun no se ha cargado el inventario" };
  }
  return { success: false, msg: "Contraseña incorrecta" };
}

// --- 3. GUARDADO DE OPERACIONES ---
function guardarTransaccion(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var f = new Date();

  if (payload.tipo === 'INVENTARIO') { // Diario
    var s = ss.getSheetByName('Respuestas');
    payload.items.forEach(function(i) {
      s.appendRow([f, payload.colaborador, payload.sucursal, payload.tipo, i.producto, i.cantidad]);
    });
    return "Inventario Diario Guardado.";
  } else if (payload.tipo === 'GENERAL') {
    var s = ss.getSheetByName('InventarioGeneral');
    payload.items.forEach(function(i) {
      s.appendRow([f, payload.sucursal, payload.colaborador, i.producto, i.proveedor, i.back, i.front, i.uso, i.total, "CARGADO"]);
    });
    return "Inventario General Guardado.";
  } else { // Entrada
    var s = ss.getSheetByName('Respuestas');
    payload.items.forEach(function(i) {
      s.appendRow([f, payload.colaborador, payload.sucursal, payload.tipo, i.producto, i.cantidad]);
    });
    return "Entrada registrada.";
  }
}

// --- 4. ADMIN DASHBOARD & CONFIG ---
function guardarConfigDias(diasArray) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('GlobalConfig');
  if (!sheet) return "Falta hoja GlobalConfig";
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 0; i < data.length; i++) { if (String(data[i][0]) === 'DIAS_GENERAL') { rowIndex = i + 1; break; } }
  var val = diasArray.join(',');
  if (rowIndex > 0) sheet.getRange(rowIndex, 2).setValue(val);
  else sheet.appendRow(['DIAS_GENERAL', val]);
  return "Días actualizados.";
}

function obtenerDashboardAdmin() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var hoy = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  var diaJava = parseInt(Utilities.formatDate(new Date(), tz, "u"));
  var diaJS = (diaJava === 7) ? 0 : diaJava;
  
  var sheetConf = ss.getSheetByName('GlobalConfig');
  var diasPermitidos = [];
  if (sheetConf && sheetConf.getLastRow() >= 2) {
    var dataC = sheetConf.getRange(2, 1, sheetConf.getLastRow()-1, 2).getValues();
    var r = dataC.find(function(row){ return String(row[0]) === 'DIAS_GENERAL'; });
    if (r) diasPermitidos = String(r[1]).split(',').map(Number);
  }
  var tocaGeneral = diasPermitidos.indexOf(diaJS) !== -1;

  var sConf = ss.getSheetByName('Config');
  var sucs = [];
  if (sConf && sConf.getLastRow() >= 2) {
    var rawSucs = sConf.getRange(2, 1, sConf.getLastRow()-1, 1).getValues();
    rawSucs.forEach(function(row){ if(row[0] && sucs.indexOf(row[0])===-1) sucs.push(row[0]); });
  }
  
  var sResp = ss.getSheetByName('Respuestas');
  var dResp = sResp && sResp.getLastRow()>=2 ? sResp.getRange(2, 1, sResp.getLastRow()-1, 4).getValues() : [];
  var okD = []; 
  dResp.forEach(function(row) { 
    if (Utilities.formatDate(new Date(row[0]), tz, "yyyy-MM-dd") == hoy && row[3] == 'INVENTARIO') okD.push(row[2]); 
  });
  
  var sGen = ss.getSheetByName('InventarioGeneral');
  var dGen = sGen && sGen.getLastRow()>=2 ? sGen.getRange(2, 1, sGen.getLastRow()-1, 2).getValues() : [];
  var okG = [];
  if (tocaGeneral) {
    dGen.forEach(function(row) { 
      if (Utilities.formatDate(new Date(row[0]), tz, "yyyy-MM-dd") == hoy) okG.push(row[1]); 
    });
  }
  
  return {
    reporte: sucs.map(function(s) {
      return {
        sucursal: s,
        diario: okD.indexOf(s) !== -1 ? "CARGADO" : "PENDIENTE",
        general: !tocaGeneral ? "NO TOCA" : (okG.indexOf(s) !== -1 ? "CARGADO" : "PENDIENTE")
      };
    }),
    configDias: diasPermitidos
  };
}

function obtenerReporteAdminDetallado() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var hoy = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  
  var sCat = ss.getSheetByName('Catalogo');
  var catMap = {};
  if (sCat && sCat.getLastRow()>=2) {
    sCat.getRange(2, 1, sCat.getLastRow()-1, 4).getValues().forEach(function(r) {
      catMap[String(r[0])] = { prov: String(r[2]||"GENERAL"), cost: Number(r[3]||0) };
    });
  }
  
  var sResp = ss.getSheetByName('Respuestas');
  if(!sResp || sResp.getLastRow()<2) return {};
  var data = sResp.getRange(2, 1, sResp.getLastRow()-1, 8).getValues();
  var rep = {};
  
  data.forEach(function(r) {
    if (Utilities.formatDate(new Date(r[0]), tz, "yyyy-MM-dd") === hoy && r[3] === 'INVENTARIO') {
      var suc=r[2], prod=r[4], cant=Number(r[5]);
      if(!rep[suc]) rep[suc] = { totalDinero: 0, grupos: {} };
      var info = catMap[prod] || { prov: "OTROS", cost: 0 };
      var totLine = cant * info.cost;
      if(!rep[suc].grupos[info.prov]) rep[suc].grupos[info.prov] = [];
      rep[suc].grupos[info.prov].push({ producto: prod, cantidad: cant, total: totLine });
      rep[suc].totalDinero += totLine;
    }
  });
  return rep;
}

function guardarCierreDia() {
  var d = obtenerDashboardAdmin().reporte;
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HistorialStatus');
  var f = new Date();
  d.forEach(function(r) { s.appendRow([f, r.sucursal, "D: " + r.diario + " | G: " + r.general]); });
  return "Cierre guardado.";
}
