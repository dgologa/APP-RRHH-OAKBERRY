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
  // Lectura explícita de Columna C (Proveedor) para Inventario Diario
  var diario = leerHoja('Catalogo', 4).map(function(r) { 
    return { 
      nombre: String(r[0]), 
      unidad: String(r[1] || "UNIDADES"), 
      proveedor: String(r[2] || "GENERAL"), // Columna C: Proveedor
      costo: Number(r[3] || 0) 
    }; 
  }).filter(function(p) { return p.nombre; });

  // NUEVO CÓDIGO INICIO: Lectura de CatalogoIngreso para Entrada de Mercancía
  var catalogoIngreso = leerHoja('CatalogoIngreso', 3).map(function(r) {
    return {
      nombre: String(r[0]),    // Col A: Producto
      proveedor: String(r[1]), // Col B: Proveedor
      unidad: String(r[2])     // Col C: Unidad
    };
  }).filter(function(p) { return p.nombre; });
  // NUEVO CÓDIGO FIN

  var general = leerHoja('CatalogoGeneral', 4).map(function(r) { 
    return { nombre: String(r[0]), unidad: String(r[1]), proveedor: String(r[2]), empaques: String(r[3]) }; 
  }).filter(function(p) { return p.nombre; });

  var empaques = leerHoja('Empaques', 3).map(function(r) { return { nombre: String(r[0]), tipo: String(r[1]), valor: Number(r[2]) }; });
  
  // Ajuste de lectura pestaña Config: Col 1 = ID, Col 2 = Nombre, Col 3 = Sucursal
  var usuarios = leerHoja('Config', 3).map(function(r) { 
    return { 
      id: String(r[0]),       // Primera columna: ID
      nombre: String(r[1]),   // Segunda columna: Nombre
      sucursal: String(r[2])  // Tercera columna: Sucursal
    }; 
  }).filter(function(u) { return u.sucursal; });

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

  // NUEVO CÓDIGO INICIO: Agregar catalogoIngreso al retorno
  return { diario: diario, general: general, ingreso: catalogoIngreso, empaques: empaques, usuarios: usuarios, admins: admins, config: { diasGeneral: diasPermitidos, diaActual: diaJS } };
  // NUEVO CÓDIGO FIN
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
    // NUEVO CÓDIGO INICIO: Reestructuración modelo datos Inventario Diario
    // Se eliminó columna D (Tipo), ahora es: Fecha, Colaborador, Sucursal, Producto, Cantidad
    payload.items.forEach(function(i) {
      s.appendRow([f, payload.colaborador, payload.sucursal, i.producto, i.cantidad]);
    });
    return "Inventario Diario Guardado.";
    // NUEVO CÓDIGO FIN

  } else if (payload.tipo === 'GENERAL') {
    var s = ss.getSheetByName('InventarioGeneral');
    payload.items.forEach(function(i) {
      s.appendRow([f, payload.sucursal, payload.colaborador, i.producto, i.proveedor, i.back, i.front, i.uso, i.total, "CARGADO"]);
    });
    return "Inventario General Guardado.";

  } else { // Entrada
    // NUEVO CÓDIGO INICIO: Guardado en IngresoMercancia
    var s = ss.getSheetByName('IngresoMercancia');
    if (!s) { // Crear si no existe (seguridad)
      s = ss.insertSheet('IngresoMercancia');
      s.appendRow(['Fecha', 'Colaborador', 'Sucursal', 'Producto', 'Cantidad']);
    }
    payload.items.forEach(function(i) {
      s.appendRow([f, payload.colaborador, payload.sucursal, i.producto, i.cantidad]);
    });
    return "Ingreso de Mercancía Registrado.";
    // NUEVO CÓDIGO FIN
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
    // Leer sucursales de Columna C (índice 2)
    var rawSucs = sConf.getRange(2, 1, sConf.getLastRow()-1, 3).getValues();
    rawSucs.forEach(function(row){ 
      var sc = row[2]; // Columna C es índice 2
      if(sc && sucs.indexOf(sc)===-1) sucs.push(sc); 
    });
  }
  
  var sResp = ss.getSheetByName('Respuestas');
  // NUEVO CÓDIGO INICIO: Ajuste lectura dashboard tras cambio de columnas en Respuestas
  // Ahora la sucursal está en la columna C (indice 2) igual que antes, 
  // pero ya no filtramos por Columna D (Tipo) porque se eliminó.
  // Asumimos que todo en Respuestas HOY es Inventario Diario a menos que se use para entradas (pero entradas va a otra hoja ahora)
  var dResp = sResp && sResp.getLastRow()>=2 ? sResp.getRange(2, 1, sResp.getLastRow()-1, 3).getValues() : [];
  var okD = []; 
  dResp.forEach(function(row) { 
    // row[0]=Fecha, row[1]=Colab, row[2]=Sucursal
    // Como Entradas va a otra hoja, todo lo que caiga aqui hoy cuenta como Inventario Diario
    if (Utilities.formatDate(new Date(row[0]), tz, "yyyy-MM-dd") == hoy) okD.push(row[2]); 
  });
  // NUEVO CÓDIGO FIN
  
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

// Función para obtener detalle específico de una sucursal y tipo
function obtenerDetalleSucursal(sucursal, tipo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var hoy = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  var resultado = [];

  if (tipo === 'DIARIO') {
    // Mapear unidades del Catalogo Diario
    var sCat = ss.getSheetByName('Catalogo');
    var mapUnidades = {};
    if(sCat && sCat.getLastRow()>=2) {
      sCat.getRange(2, 1, sCat.getLastRow()-1, 2).getValues().forEach(r => mapUnidades[String(r[0])] = String(r[1]));
    }
    
    var sResp = ss.getSheetByName('Respuestas');
    if(sResp && sResp.getLastRow()>=2){
      // NUEVO CÓDIGO INICIO: Ajuste lectura detalle tras eliminación de columna D
      // Estructura: A=Fecha, B=Colab, C=Sucursal, D=Producto, E=Cantidad
      var data = sResp.getRange(2, 1, sResp.getLastRow()-1, 5).getValues();
      data.forEach(r => {
        if (Utilities.formatDate(new Date(r[0]), tz, "yyyy-MM-dd") === hoy && r[2] === sucursal) {
          resultado.push({ producto: r[3], cantidad: r[4], unidad: mapUnidades[r[3]] || 'N/A' });
        }
      });
      // NUEVO CÓDIGO FIN
    }
  } else if (tipo === 'GENERAL') {
    // Mapear unidades del Catalogo General
    var sCatG = ss.getSheetByName('CatalogoGeneral');
    var mapUnidadesG = {};
    if(sCatG && sCatG.getLastRow()>=2) {
      sCatG.getRange(2, 1, sCatG.getLastRow()-1, 2).getValues().forEach(r => mapUnidadesG[String(r[0])] = String(r[1]));
    }

    var sInvG = ss.getSheetByName('InventarioGeneral');
    if(sInvG && sInvG.getLastRow()>=2){
      var data = sInvG.getRange(2, 1, sInvG.getLastRow()-1, 9).getValues(); // Hasta columna total
      data.forEach(r => {
        // Col 0:Fecha, Col 1:Sucursal, Col 3:Producto, Col 8:Total
        if (Utilities.formatDate(new Date(r[0]), tz, "yyyy-MM-dd") === hoy && r[1] === sucursal) {
          resultado.push({ producto: r[3], cantidad: r[8], unidad: mapUnidadesG[r[3]] || 'N/A' });
        }
      });
    }
  }
  return resultado;
}

// Función de Cierre Automático para ejecutar con Trigger a las 23:59
function procesarCierreAutomatico() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var hoy = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  
  // 1. Obtener sucursales únicas de Config Columna C
  var sConf = ss.getSheetByName('Config');
  var sucs = [];
  if (sConf && sConf.getLastRow() >= 2) {
    var rawSucs = sConf.getRange(2, 1, sConf.getLastRow()-1, 3).getValues();
    rawSucs.forEach(function(row){ if(row[2] && sucs.indexOf(row[2])===-1) sucs.push(row[2]); });
  }

  // 2. Verificar Configuración de Días para General
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

  // 3. Checar status en Respuestas (Diario)
  var sResp = ss.getSheetByName('Respuestas');
  var checkDiario = [];
  if(sResp && sResp.getLastRow()>=2) {
    // NUEVO CÓDIGO INICIO: Ajuste lectura cierre tras eliminación columna D
    // Validamos solo por fecha y sucursal en col C
    sResp.getRange(2, 1, sResp.getLastRow()-1, 3).getValues().forEach(r => {
      if(Utilities.formatDate(new Date(r[0]), tz, "yyyy-MM-dd") == hoy) checkDiario.push(r[2]);
    });
    // NUEVO CÓDIGO FIN
  }

  // 4. Checar status en InventarioGeneral
  var sGen = ss.getSheetByName('InventarioGeneral');
  var checkGeneral = [];
  if(sGen && sGen.getLastRow()>=2 && tocaGeneral) {
    sGen.getRange(2, 1, sGen.getLastRow()-1, 2).getValues().forEach(r => {
      if(Utilities.formatDate(new Date(r[0]), tz, "yyyy-MM-dd") == hoy) checkGeneral.push(r[1]);
    });
  }

  // 5. Escribir en Envio_Inventarios
  var sEnvio = ss.getSheetByName('Envio_Inventarios');
  if (!sEnvio) {
    sEnvio = ss.insertSheet('Envio_Inventarios');
    sEnvio.appendRow(['Fecha', 'Sucursal', 'Validacion_Diario', 'Validacion_General']);
  }
  
  var timestamp = new Date();
  sucs.forEach(suc => {
    var valDiario = checkDiario.includes(suc) ? 1 : 0;
    var valGeneral = "nada";
    if (tocaGeneral) {
      valGeneral = checkGeneral.includes(suc) ? 1 : 0;
    }
    sEnvio.appendRow([timestamp, suc, valDiario, valGeneral]);
  });
}
