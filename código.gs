/**
 * CONTROL DE PEDIDOS - LEVAN
 * Backend con soporte para Fecha de Entrega (Columna L) y Reporte por Periodo.
 */

const CONFIG = {
  SHEET_PEDIDOS: "BD_Pedidos",
  SHEET_PRECIOS: "Lista_Precios"
};

const LISTA_PRECIOS_DEFAULT = {
  "Hogaza Natural": 125,
  "Hogaza Arandano": 175,
  "Hogaza Integral": 150,
  "Hogaza ½ Datil/Pistache": 120,
  "Hogaza ½ Natural": 90,
  "Hogaza ½ Arandano y Nuez": 120,
  "Hogaza ½ Integral": 100,
  "Concha Vainilla": 20,
  "Concha Chocolate": 20,
  "Barra Pan Caja": 150,
  "Pan Hamburguesa 5 PZ": 100,
  "Pan de Muerto": 40,
  "Rosca": 175,
  "Pq 25 pz Pan Mesa": 170,
  "Pq 35 pz Pan Mesa": 230
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Control de Pedidos - LEVAN')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function normalizarEstatus(val) {
  if (!val) return "Pendiente";
  const str = String(val).trim().toLowerCase();
  if (str.includes("proceso")) return "En Proceso";
  if (str.includes("listo") || str.includes("hecho")) return "Listo";
  if (str.includes("entregado")) return "Entregado";
  return "Pendiente";
}

function formatearFechaSegura(val, formato = "yyyy-MM-dd HH:mm") {
  if (!val) return "";
  if (val instanceof Date && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), formato);
  }
  try {
    const d = new Date(val);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), formato);
    }
  } catch (e) {}
  return String(val).trim();
}

/**
 * Carga inicial de datos desde la hoja BD_Pedidos (Columnas A - L).
 */
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let listaPrecios = { ...LISTA_PRECIOS_DEFAULT };
  const sheetPrecios = ss.getSheetByName(CONFIG.SHEET_PRECIOS);
  if (sheetPrecios && sheetPrecios.getLastRow() > 1) {
    const dataPrecios = sheetPrecios.getRange(2, 1, sheetPrecios.getLastRow() - 1, 2).getValues();
    dataPrecios.forEach(([prod, precio]) => {
      if (prod && !isNaN(parseFloat(precio))) {
        listaPrecios[String(prod).trim()] = Number(precio);
      }
    });
  }

  let sheetPedidos = ss.getSheetByName(CONFIG.SHEET_PEDIDOS);
  if (!sheetPedidos) {
    sheetPedidos = ss.insertSheet(CONFIG.SHEET_PEDIDOS);
    sheetPedidos.appendRow(["Fecha Solicitud", "No Pedido", "Vendedor", "Cliente", "Producto", "Cantidad", "Precio", "Comentarios", "Estatus", "¿Pagado?", "Tipo de Pago", "Fecha Entrega"]);
    return { pedidos: [], listaPrecios };
  }

  const lastRow = sheetPedidos.getLastRow();
  if (lastRow <= 1) {
    return { pedidos: [], listaPrecios };
  }

  const values = sheetPedidos.getRange(2, 1, lastRow - 1, 12).getValues();
  const pedidos = [];
  
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rawNoPedido = row[1];
    if (rawNoPedido === "" || rawNoPedido === null || rawNoPedido === undefined) continue;

    const noPedidoNum = parseInt(rawNoPedido, 10);
    if (isNaN(noPedidoNum)) continue;

    pedidos.push({
      fechaSolicitud: formatearFechaSegura(row[0], "yyyy-MM-dd HH:mm"),
      noPedido: noPedidoNum,
      vendedor: String(row[2] || "").trim(),
      cliente: String(row[3] || "").trim(),
      producto: String(row[4] || "").trim(),
      cantidad: parseFloat(row[5]) || 0,
      precio: parseFloat(row[6]) || 0,
      comentarios: String(row[7] || "").trim(),
      estatus: normalizarEstatus(row[8]),
      pagado: String(row[9] || "").trim(),
      tipoPago: String(row[10] || "").trim(),
      fechaEntrega: formatearFechaSegura(row[11], "yyyy-MM-dd")
    });
  }

  return { pedidos, listaPrecios };
}

function getNextOrderNumber(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1000;
  const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const maxNo = values.reduce((max, row) => {
    const val = parseInt(row[0], 10);
    return (!isNaN(val) && val > max) ? val : max;
  }, 999);
  return maxNo + 1;
}

function guardarPedido(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    throw new Error("El servidor está ocupado. Intente de nuevo.");
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_PEDIDOS);
    const now = new Date();
    let noPedido = payload.noPedido;

    if (!noPedido) {
      noPedido = getNextOrderNumber(sheet);
    } else {
      eliminarItemsPorPedido(sheet, noPedido);
    }

    const newRows = payload.items.map(item => [
      now,
      noPedido,
      payload.vendedor,
      payload.cliente,
      item.producto,
      item.cantidad,
      item.precio,
      payload.comentarios || "",
      payload.estatus || "Pendiente",
      payload.pagado ? "Si" : "No",
      payload.tipoPago || "",
      payload.fechaEntrega || ""
    ]);

    if (newRows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, newRows.length, 12).setValues(newRows);
    }

    return { success: true, noPedido };
  } finally {
    lock.releaseLock();
  }
}

function eliminarItemsPorPedido(sheet, noPedido) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (parseInt(data[i][0], 10) === parseInt(noPedido, 10)) {
      sheet.deleteRow(i + 2);
    }
  }
}

/**
 * Obtiene el resumen de requerimientos consolidado y la lista de comandas agrupadas por día.
 * @param {string} fechaInicioStr - Fecha en formato 'YYYY-MM-DD'
 * @param {string} fechaFinStr - Fecha en formato 'YYYY-MM-DD'
 * @returns {Object} { dias, matriz, comandasPorDia }
 */
function getRequerimientosPorPeriodo(fechaInicioStr, fechaFinStr) {
  if (!fechaInicioStr || !fechaFinStr) {
    throw new Error("Debe seleccionar las fechas de inicio y fin.");
  }

  const { pedidos } = getInitialData();
  const inicio = new Date(fechaInicioStr + "T00:00:00");
  const fin = new Date(fechaFinStr + "T23:59:59");

  const estatusValidos = ["Pendiente", "En Proceso", "Listo"];
  const pedidosFiltrados = (pedidos || []).filter(p => {
    if (!p.fechaEntrega || !estatusValidos.includes(p.estatus)) return false;
    const fEntrega = new Date(p.fechaEntrega + "T00:00:00");
    return fEntrega >= inicio && fEntrega <= fin;
  });

  // Generar lista continua de días (YYYY-MM-DD)
  const dias = [];
  let curr = new Date(inicio);
  while (curr <= fin) {
    dias.push(Utilities.formatDate(curr, Session.getScriptTimeZone(), "yyyy-MM-dd"));
    curr.setDate(curr.getDate() + 1);
  }

  // 1. Matriz de Requerimientos por Producto
  const matriz = {};
  
  // 2. Mapa de Pedidos/Comandas Agrupados por Fecha de Entrega
  const comandasMapa = {};
  dias.forEach(d => comandasMapa[d] = {});

  pedidosFiltrados.forEach(p => {
    const fDia = p.fechaEntrega;
    const prod = p.producto;
    const cant = Number(p.cantidad) || 0;

    // Matriz de productos
    if (!matriz[prod]) {
      matriz[prod] = {};
      dias.forEach(d => matriz[prod][d] = 0);
    }
    if (matriz[prod][fDia] !== undefined) {
      matriz[prod][fDia] += cant;
    }

    // Agrupamiento por comanda/pedido
    if (comandasMapa[fDia]) {
      const key = p.noPedido;
      if (!comandasMapa[fDia][key]) {
        comandasMapa[fDia][key] = {
          noPedido: p.noPedido,
          vendedor: p.vendedor,
          cliente: p.cliente,
          comentarios: p.comentarios,
          estatus: p.estatus,
          pagado: p.pagado,
          tipoPago: p.tipoPago,
          fechaSolicitud: p.fechaSolicitud,
          fechaEntrega: p.fechaEntrega,
          items: []
        };
      }
      comandasMapa[fDia][key].items.push({
        producto: p.producto,
        cantidad: p.cantidad,
        precio: p.precio,
        total: p.cantidad * p.precio
      });
    }
  });

  // Convertir mapas de comandas a arrays por día
  const comandasPorDia = {};
  dias.forEach(d => {
    comandasPorDia[d] = Object.values(comandasMapa[d] || {});
  });

  return { dias, matriz, comandasPorDia };
}
