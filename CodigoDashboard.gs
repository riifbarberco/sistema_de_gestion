/**
 * RIIF — BLOQUE DASHBOARD
 *
 * PEGA ESTE BLOQUE COMPLETO AL FINAL DE Codigo.gs (después de la última
 * llave de quitarDisparadores). No toques nada de lo que ya existe.
 *
 * QUÉ AGREGA
 *  · Ruta ?p=dashboard en doGet → sirve la página Dashboard
 *  · getDatosDashboard(barbero, desde, hasta) → métricas para el dashboard
 *
 * CAMBIO EN doGet (único lugar donde hay que editar el código existente):
 *   Busca la función doGet al inicio del archivo y reemplázala
 *   por la versión de abajo.
 */

// ================================================================
//  REEMPLAZA la función doGet existente por esta (es la única
//  modificación al código viejo — todo lo demás va al final).
// ================================================================
//
// function doGet(e) {
//   var p  = e && e.parameter ? e.parameter.p : null;
//   var id = e && e.parameter ? e.parameter.c : null;
//   var tpl;
//   if (p === 'dashboard') {
//     tpl = HtmlService.createTemplateFromFile('Dashboard');
//   } else if (id) {
//     tpl = HtmlService.createTemplateFromFile('Cancelar');
//     tpl.reservaId = id;
//   } else {
//     tpl = HtmlService.createTemplateFromFile('Reservar');
//     tpl.reservaId = '';
//   }
//   return tpl.evaluate()
//     .setTitle('RIIF Barber Co.')
//     .addMetaTag('viewport', 'width=device-width, initial-scale=1')
//     .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
// }
//
// ================================================================
//  FUNCIONES NUEVAS — pega desde aquí hasta el final del archivo
// ================================================================

/**
 * Devuelve las métricas del Registro para un barbero y un rango de fechas.
 *
 * @param {string} barbero   - Nombre exacto del barbero (como aparece en Config)
 * @param {string} desde     - Fecha de inicio en formato YYYY-MM-DD
 * @param {string} hasta     - Fecha de fin en formato YYYY-MM-DD
 * @returns {Object} objeto con totalServicios, totalComision, totalPropinas,
 *                   totalCobrado, porPeriodo[], topServicios[], servicios[]
 */
function getDatosDashboard(barbero, desde, hasta, verTodo) {
  var hg = libro().getSheetByName('Registro');
  if (!hg || hg.getLastRow() < 2) {
    return _dashboardVacio();
  }

  // Leer todo el Registro. Se piden hasta 28 columnas (productos 22-27 y su
  // valor en la 28) pero se respeta el ancho real de la hoja para no salirse
  // de la cuadrícula.
  var anchoLeer = Math.max(16, Math.min(COL_VALOR_PROD, hg.getLastColumn()));
  var datos = hg.getRange(2, 1, hg.getLastRow() - 1, anchoLeer).getValues();

  var servicios     = [];
  var totalCobrado  = 0;
  var totalComision = 0;
  var totalPropinas = 0;
  var totalBarberia = 0;    // Neto barbería: servicio − comisión + productos
  var totalProductos = 0;   // $ de productos vendidos
  var precios = null;       // precios del Inventario, solo si alguna fila los necesita
  var contPorDia    = {};   // { 'YYYY-MM-DD': { count, total } }
  var contServicios = {};   // { 'Nombre servicio': { count, total } }

  datos.forEach(function(f) {
    // Columnas del Registro (base-1 en Sheets → base-0 aquí):
    //  0:Fecha  1:Hora  2:Barbero  3:Cliente  4:Teléfono  5:Servicio
    //  6:Valor  7:Desc  8:Total    9:Propina  10:Método   11:Estado
    // 12:Comisión  13:PagoBarbero  14:Neto  15:Notas
    // 16:Año  17:Mes  18:DiaSem  19:HoraBloque  20:Duración
    // 21:Productos (texto)  27:Valor productos

    var fBarbero = String(f[2] || '').trim();
    if (!verTodo && fBarbero !== barbero) return;

    var estado = String(f[11] || '').trim();
    if (estado === 'Cancelado') return; // ignorar canceladas

    var iso = _isoCeldaRegistro_(f[0]);
    if (!iso) return; // fecha inválida, saltar fila

    if (iso < desde || iso > hasta) return;

    var total       = Number(f[8])  || 0;
    var descuento   = Number(f[7])  || 0;
    var propina     = Number(f[9])  || 0;
    var comision    = Number(f[12]) || 0;
    var pagoBarbero = Number(f[13]) || 0;
    var valorProd   = Number(f[27]) || 0;
    var nombreSrv   = String(f[5] || '').trim();
    var hora        = String(f[1] || '').trim();
    var cliente     = String(f[3] || '').trim();
    var metodo      = String(f[10] || '').trim();

    // id de la cita, guardado en la col 16 (Notas) como "Cita <id>"
    var nota = String(f[15] || '').trim();
    var citaId = (nota.indexOf('Cita ') === 0) ? nota.substring(5).trim() : '';

    // Productos: { 'Agua': {v,r}, ... } — acepta formato nuevo (texto) y viejo.
    var productos = _leerProductosRegistro_([f[21], f[22], f[23], f[24], f[25], f[26]]);

    // Filas cerradas antes de existir la col "Valor productos": se calcula con
    // el precio actual para que las bebidas igual cuenten.
    if (!valorProd && Object.keys(productos).length && (f[27] === '' || f[27] == null)) {
      if (!precios) precios = _preciosProductos_();
      valorProd = _valorProductos_(productos, precios);
    }

    // Pago barbería = servicio − comisión + bebidas vendidas (100% barbería).
    // Se calcula aquí y no se lee de la col O, para que no dependa de que la
    // fórmula de la hoja esté al día.
    var atendido = (estado === 'Atendido');
    var neto = atendido ? total - comision + valorProd : 0;
    if (!atendido) valorProd = 0;

    // Total cobrado = servicio (con descuento) + productos vendidos. Así
    // Total cobrado + Propinas = Pago barberos + Pago barbería.
    totalCobrado   += total + valorProd;
    totalComision  += pagoBarbero;   // "pago barbero" = comisión + propina
    totalPropinas  += propina;
    totalBarberia  += neto;
    totalProductos += valorProd;

    // Acumular por día
    if (!contPorDia[iso]) contPorDia[iso] = { count: 0, total: 0 };
    contPorDia[iso].count++;
    contPorDia[iso].total += pagoBarbero;

    // Acumular por nombre de servicio
    if (nombreSrv) {
      if (!contServicios[nombreSrv]) contServicios[nombreSrv] = { count: 0, total: 0 };
      contServicios[nombreSrv].count++;
      contServicios[nombreSrv].total += pagoBarbero;
    }

    // Agregar a la lista de servicios individuales
    servicios.push({
      id: citaId,
      fecha: iso, hora: hora, barbero: fBarbero,
      servicio: nombreSrv, cliente: cliente,
      total: total, pagoBarbero: pagoBarbero, neto: neto, valorProductos: valorProd,
      propina: propina, metodo: metodo, descuento: descuento,
      productos: productos,
      estado: estado || 'Pendiente'
    });
  });

  // Ordenar servicios más recientes primero
  servicios.sort(function(a, b) {
    if (b.fecha !== a.fecha) return b.fecha > a.fecha ? 1 : -1;
    return b.hora > a.hora ? 1 : -1;
  });

  // Construir array porPeriodo para el gráfico
  var porPeriodo = _construirPeriodo(desde, hasta, contPorDia);

  // Top servicios (máx. 8)
  var topServicios = Object.keys(contServicios)
    .map(function(k) { return { nombre: k, count: contServicios[k].count, total: contServicios[k].total }; })
    .sort(function(a, b) { return b.count - a.count; })
    .slice(0, 8);

  return {
    totalServicios: servicios.length,
    totalCobrado:   Math.round(totalCobrado),
    // Sin redondear al peso entero: con comisiones al 50% (u otro % impar)
    // una sola cita ya deja décimas (ej. 12.5), y Math.round() las subía a 13.
    // Se limpia a 2 decimales solo para evitar residuos de coma flotante
    // (ej. 14.800000000000001), sin perder la parte exacta.
    totalComision:  Math.round(totalComision * 100) / 100,
    totalPropinas:  Math.round(totalPropinas),
    totalBarberia:  Math.round(totalBarberia * 100) / 100,
    totalProductos: Math.round(totalProductos),
    porPeriodo:     porPeriodo,
    topServicios:   topServicios,
    servicios:      servicios
  };
}

/** Fecha ISO de la columna Fecha del Registro (Date, DD/MM/YYYY o YYYY-MM-DD); '' si no se entiende. */
function _isoCeldaRegistro_(v) {
  if (v instanceof Date) return aISO(v);
  var s = String(v || '').trim();
  var mDMY = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  var mYMD = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (mDMY) return mDMY[3] + '-' + ('0' + mDMY[2]).slice(-2) + '-' + ('0' + mDMY[1]).slice(-2);
  if (mYMD) return mYMD[1] + '-' + ('0' + mYMD[2]).slice(-2) + '-' + ('0' + mYMD[3]).slice(-2);
  return '';
}

// ================================================================
//  ANÁLISIS — gráficas para tomar decisiones
//  Dueño: toda la barbería. Barbero: solo sus propios servicios.
// ================================================================

/**
 * Agregados del Registro entre `desde` y `hasta` (yyyy-mm-dd), más los
 * totales del período anterior del mismo largo para comparar.
 * Ingresos = servicio cobrado (con descuento) + bebidas vendidas, solo "Atendido".
 *
 * @returns { ok, desde, hasta, actual:{...}, anterior:{...}, porFecha, heat,
 *            porHora, porDow, dowOcurrencias, porBarbero, porServicio, bebidas }
 */
function getAnalisisDashboard(token, desde, hasta) {
  var s;
  try { s = _sesion(token); } catch (e) { return { ok: false, error: e.message }; }
  var soloBarbero = (s.rol === 'Dueño') ? '' : s.nombre;
  desde = String(desde || '').slice(0, 10);
  hasta = String(hasta || '').slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(desde) || !/^\d{4}-\d{2}-\d{2}$/.test(hasta) || desde > hasta) {
    return { ok: false, error: 'Rango de fechas no válido.' };
  }

  // Período anterior del mismo largo, justo antes de `desde`
  var d0 = new Date(desde + 'T12:00:00'), d1 = new Date(hasta + 'T12:00:00');
  var dias = Math.round((d1 - d0) / 86400000) + 1;
  var prevHasta = Utilities.formatDate(new Date(d0.getTime() - 86400000), TZ, 'yyyy-MM-dd');
  var prevDesde = Utilities.formatDate(new Date(d0.getTime() - dias * 86400000), TZ, 'yyyy-MM-dd');

  function totalesVacios() {
    return { atendidos: 0, noShow: 0, cancelados: 0, ingresos: 0, ingresoServicios: 0,
             ingresoBebidas: 0, barberia: 0, propinas: 0, pagoBarbero: 0 };
  }
  var out = {
    ok: true, desde: desde, hasta: hasta, prevDesde: prevDesde, prevHasta: prevHasta,
    barbero: soloBarbero,   // '' = toda la barbería
    actual: totalesVacios(), anterior: totalesVacios(),
    porFecha: {},          // { iso: { n, ingresos } }
    heat: [],              // heat[dow 0=Lun][hora 0-23] = servicios atendidos
    porBarbero: {},        // { nombre: { n, ingresos } }
    porServicio: {},       // { nombre: n } (los combos se cuentan por separado)
    bebidas: {},           // { nombre: { v, r } }
    dowOcurrencias: [0, 0, 0, 0, 0, 0, 0]   // cuántos lunes, martes… tiene el rango
  };
  for (var k = 0; k < 7; k++) { var fila = []; for (var h = 0; h < 24; h++) fila.push(0); out.heat.push(fila); }
  for (var t = 0; t < dias; t++) {
    out.dowOcurrencias[(new Date(d0.getTime() + t * 86400000).getDay() + 6) % 7]++;
  }

  var hg = libro().getSheetByName('Registro');
  if (!hg || hg.getLastRow() < 2) return out;
  var ancho = Math.max(16, Math.min(COL_VALOR_PROD, hg.getLastColumn()));
  var datos = hg.getRange(2, 1, hg.getLastRow() - 1, ancho).getValues();
  var precios = null;

  datos.forEach(function (f) {
    var iso = _isoCeldaRegistro_(f[0]);
    if (!iso || iso < prevDesde || iso > hasta) return;
    if (soloBarbero && String(f[2] || '').trim() !== soloBarbero) return;
    var esActual = iso >= desde;
    var tot = esActual ? out.actual : out.anterior;

    var estado = String(f[11] || '').trim();
    if (estado === 'Cancelado') { tot.cancelados++; return; }
    if (estado === 'No asistió') { tot.noShow++; return; }
    if (estado !== 'Atendido') return;   // pendientes no cuentan todavía

    var total = Number(f[8]) || 0;
    var comision = Number(f[12]) || 0;
    var productos = _leerProductosRegistro_([f[21], f[22], f[23], f[24], f[25], f[26]]);
    var valorProd = Number(f[27]) || 0;
    if (!valorProd && Object.keys(productos).length && (f[27] === '' || f[27] == null)) {
      if (!precios) precios = _preciosProductos_();
      valorProd = _valorProductos_(productos, precios);
    }
    var ingreso = total + valorProd;

    tot.atendidos++;
    tot.ingresos += ingreso;
    tot.ingresoServicios += total;
    tot.ingresoBebidas += valorProd;
    tot.barberia += total - comision + valorProd;
    tot.propinas += Number(f[9]) || 0;
    tot.pagoBarbero += Number(f[13]) || 0;
    if (!esActual) return;

    var pf = out.porFecha[iso] || (out.porFecha[iso] = { n: 0, ingresos: 0 });
    pf.n++; pf.ingresos += ingreso;

    var dow = (new Date(iso + 'T12:00:00').getDay() + 6) % 7;
    var hhmm = aHHMM(f[1]);
    var hora = parseInt(String(hhmm).split(':')[0], 10);
    if (!isNaN(hora) && hora >= 0 && hora < 24) out.heat[dow][hora]++;

    var barbero = String(f[2] || '').trim() || '—';
    var pb = out.porBarbero[barbero] || (out.porBarbero[barbero] = { n: 0, ingresos: 0 });
    pb.n++; pb.ingresos += ingreso;

    String(f[5] || '').split(SEPARADOR_COMBO).forEach(function (n) {
      n = n.trim();
      if (n) out.porServicio[n] = (out.porServicio[n] || 0) + 1;
    });

    Object.keys(productos).forEach(function (n) {
      var b = out.bebidas[n] || (out.bebidas[n] = { v: 0, r: 0 });
      b.v += Number(productos[n].v) || 0;
      b.r += Number(productos[n].r) || 0;
    });
  });
  return out;
}

/** Resultado vacío cuando no hay datos */
function _dashboardVacio() {
  return {
    totalServicios: 0, totalCobrado: 0,
    totalComision: 0,  totalPropinas: 0,
    totalBarberia: 0,  totalProductos: 0,
    porPeriodo: [], topServicios: [], servicios: []
  };
}

/**
 * Construye el array de puntos para el gráfico.
 * Si el rango es ≤ 31 días → agrupa por día.
 * Si el rango es ≤ 92 días → agrupa por semana.
 * Si el rango es > 92 días → agrupa por mes.
 */
function _construirPeriodo(desde, hasta, contPorDia) {
  // Diferencia en días. OJO: se parsea como hora LOCAL ('...T00:00:00'), no UTC,
  // para que _isoDate (que lee getFullYear/getMonth/getDate locales) coincida
  // con las claves de contPorDia (generadas con aISO / zona de la hoja).
  var d1 = new Date(desde + 'T00:00:00'), d2 = new Date(hasta + 'T00:00:00');
  var diff = Math.round((d2 - d1) / 86400000) + 1;

  var meses = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  var dias  = ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'];

  if (diff <= 1) {
    // Un solo día → un único punto con el total del día
    var item = contPorDia[desde] || { count: 0, total: 0 };
    return [{ label: fmtLabelFecha(desde), count: item.count, total: item.total }];
  }

  if (diff <= 31) {
    // Por día
    var result = [];
    var cur = new Date(d1);
    while (cur <= d2) {
      var iso = _isoDate(cur);
      var item2 = contPorDia[iso] || { count: 0, total: 0 };
      result.push({ label: dias[cur.getDay()] + ' ' + cur.getDate(), count: item2.count, total: item2.total });
      cur.setDate(cur.getDate() + 1);
    }
    // Filtrar días sin datos si hay muchos días
    if (diff > 14) {
      result = result.filter(function(r) { return r.count > 0; });
    }
    return result;
  }

  if (diff <= 366) {
    // Por mes
    var porMes = {};
    Object.keys(contPorDia).forEach(function(iso) {
      var ym = iso.slice(0, 7);
      if (!porMes[ym]) porMes[ym] = { count: 0, total: 0 };
      porMes[ym].count += contPorDia[iso].count;
      porMes[ym].total += contPorDia[iso].total;
    });
    return Object.keys(porMes).sort().map(function(ym) {
      var p = ym.split('-');
      return { label: meses[parseInt(p[1], 10) - 1] + ' ' + p[0], count: porMes[ym].count, total: porMes[ym].total };
    });
  }

  // Más de un año → por mes también
  var porMes2 = {};
  Object.keys(contPorDia).forEach(function(iso) {
    var ym = iso.slice(0, 7);
    if (!porMes2[ym]) porMes2[ym] = { count: 0, total: 0 };
    porMes2[ym].count += contPorDia[iso].count;
    porMes2[ym].total += contPorDia[iso].total;
  });
  return Object.keys(porMes2).sort().map(function(ym) {
    var p = ym.split('-');
    return { label: meses[parseInt(p[1], 10) - 1] + ' ' + p[0], count: porMes2[ym].count, total: porMes2[ym].total };
  });
}

function fmtLabelFecha(iso) {
  var meses = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
  var p = iso.split('-');
  return parseInt(p[2], 10) + ' ' + meses[parseInt(p[1], 10) - 1];
}

function _isoDate(d) {
  return d.getFullYear() + '-' +
    ('0' + (d.getMonth() + 1)).slice(-2) + '-' +
    ('0' + d.getDate()).slice(-2);
}

// ================================================================
//  FASE B — Autenticación por PIN y cierre de servicios
// ================================================================

/**
 * Secreto para firmar los tokens de sesión. Vive en las propiedades del script,
 * nunca en el código ni en la hoja. Se genera solo la primera vez.
 */
function _dashSecreto_() {
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('DASH_SECRET');
  if (!s) {
    s = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('DASH_SECRET', s);
  }
  return s;
}

function _dashFirma_(txt) {
  var raw = Utilities.computeHmacSha256Signature(txt, _dashSecreto_());
  return Utilities.base64EncodeWebSafe(raw);
}

/**
 * Valida un token de sesión. Devuelve { nombre, rol } o lanza error.
 * Formato del token:  base64url(JSON {n,r,e}) + '.' + base64url(HMAC)
 */
function _sesion(token) {
  if (!token || String(token).indexOf('.') < 0) throw new Error('Sesión no válida. Vuelve a entrar con tu PIN.');
  var partes = String(token).split('.');
  var cuerpo = partes[0], firma = partes[1];
  if (_dashFirma_(cuerpo) !== firma) throw new Error('Sesión no válida. Vuelve a entrar con tu PIN.');
  var datos;
  try {
    datos = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(cuerpo)).getDataAsString());
  } catch (e) {
    throw new Error('Sesión no válida. Vuelve a entrar con tu PIN.');
  }
  if (!datos || !datos.e || Date.now() > datos.e) throw new Error('Tu sesión venció. Vuelve a entrar con tu PIN.');

  // El ROL se re-lee de la hoja Barberos (si le cambian el Rol a alguien,
  // no tiene que volver a entrar). Si la hoja falla, se usa el del token.
  var rol = datos.r;
  try {
    var b = leerBarberos()[datos.n];
    if (b) rol = _rolDe_(b.rol);
  } catch (e) {}
  return { nombre: datos.n, rol: rol };
}

/**
 * Login del dashboard. Compara el PIN contra la hoja Barberos.
 * @returns {Object} { ok, rol, nombre, token } o { ok:false, error }
 */
function iniciarSesion(nombre, pin) {
  nombre = String(nombre || '').trim();
  pin    = String(pin || '').replace(/\D/g, '');
  if (!nombre) return { ok: false, error: 'Elige tu nombre.' };
  if (pin.length < 4) return { ok: false, error: 'El PIN son 4 dígitos.' };

  var b = leerBarberos()[nombre];
  if (!b) return { ok: false, error: 'No encontramos "' + nombre + '" en la hoja Barberos.' };
  if (!b.pin) return { ok: false, error: 'Tu usuario todavía no tiene PIN. Pídeselo al administrador.' };
  if (b.pin !== pin) return { ok: false, error: 'PIN incorrecto.' };

  var rol = _rolDe_(b.rol);
  var cuerpo = Utilities.base64EncodeWebSafe(JSON.stringify({
    n: nombre, r: rol, e: Date.now() + 30 * 86400000
  }));
  return { ok: true, rol: rol, nombre: nombre, token: cuerpo + '.' + _dashFirma_(cuerpo) };
}

/** Normaliza el texto de la columna Rol a 'Dueño' o 'barbero'. */
function _rolDe_(texto) {
  return /due|admin|jefe|propietar|owner/i.test(String(texto || '')) ? 'Dueño' : 'barbero';
}

/** Perfil actual (identidad + rol fresco). El rol ya lo re-lee _sesion. */
function getPerfilDashboard(token) {
  return _sesion(token);
}

/**
 * Cambia el PIN del usuario que inició sesión. Pide el PIN actual (no alcanza
 * con el token) para que, si alguien encuentra el celular ya logueado, no
 * pueda cambiarle el PIN a otra persona sin saber el que tiene hoy.
 *
 * @param {string} token
 * @param {string} pinActual
 * @param {string} pinNuevo   - 4 dígitos
 * @returns {Object} { ok } o { ok:false, error }
 */
function cambiarPinDashboard(token, pinActual, pinNuevo) {
  var s;
  try { s = _sesion(token); } catch (e) { return { ok: false, error: e.message }; }

  pinActual = String(pinActual || '').replace(/\D/g, '');
  pinNuevo  = String(pinNuevo || '').replace(/\D/g, '');
  if (pinNuevo.length !== 4) return { ok: false, error: 'El PIN nuevo debe tener 4 dígitos.' };
  if (pinNuevo === pinActual) return { ok: false, error: 'Elige un PIN distinto al que ya tienes.' };

  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch (e) { return { ok: false, error: 'El sistema está ocupado. Intenta de nuevo.' }; }

  try {
    var b = leerBarberos()[s.nombre];
    if (!b) return { ok: false, error: 'No encontramos tu usuario en la hoja Barberos.' };
    if (b.colPin < 1) return { ok: false, error: 'La hoja Barberos no tiene columna PIN. Avisa al administrador.' };
    if (!b.pin || b.pin !== pinActual) return { ok: false, error: 'El PIN actual no es correcto.' };

    var h = libro().getSheetByName(HOJA_BARBEROS);
    h.getRange(b.fila, b.colPin).setNumberFormat('@').setValue(pinNuevo);
    return { ok: true };

  } catch (err) {
    return { ok: false, error: 'No se pudo cambiar el PIN: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Nombres que aparecen en el selector de login del Dashboard:
 * los barberos de Config + cualquier fila de la hoja Barberos que ya tenga PIN
 * (así aparece también la cuenta de administrador, aunque no corte pelo).
 */
function getUsuariosDashboard() {
  var lista = [];
  var vistos = {};

  try {
    leerConfig().barberos.forEach(function (b) {
      if (b.nombre && !vistos[b.nombre]) { vistos[b.nombre] = true; lista.push(b.nombre); }
    });
  } catch (e) { /* si Config falla, seguimos con lo que haya en Barberos */ }

  try {
    var bs = leerBarberos();
    Object.keys(bs).forEach(function (nombre) {
      if (vistos[nombre]) return;
      if (nombre === CALENDARIO_GENERAL) return;      // fila del calendario general
      if (bs[nombre] && bs[nombre].pin) { vistos[nombre] = true; lista.push(nombre); }
    });
  } catch (e) {}

  return { usuarios: lista };
}

/**
 * Métodos de pago, estados y productos para la pantalla de cierre.
 * NUNCA lanza: si algo falla, devuelve valores por defecto para que el
 * dashboard siempre pueda abrir el modal de cierre.
 */
function getOpcionesCierre() {
  var metodos = [], estados = [], productos = [];
  try {
    var h = libro().getSheetByName(HOJA_CONFIG);
    if (h) {
      h.getRange('H5:H14').getValues().forEach(function (f) {
        if (String(f[0]).trim()) metodos.push(String(f[0]).trim());
      });
      h.getRange('J5:J10').getValues().forEach(function (f) {
        if (String(f[0]).trim()) estados.push(String(f[0]).trim());
      });
    }
  } catch (e) {}
  try { productos = _productosCatalogo_(); } catch (e) { productos = []; }

  // Catálogo de servicios: el modal de cierre deja ajustar lo que de verdad se hizo.
  var servicios = [];
  try { servicios = leerConfig().servicios || []; } catch (e) { servicios = []; }

  if (!metodos.length) metodos = ['Efectivo', 'Nequi', 'Daviplata', 'Transferencia', 'Tarjeta'];
  if (!estados.length) estados = ['Atendido', 'No asistió', 'Cancelado'];

  return { metodos: metodos, estados: estados, productos: productos, servicios: servicios };
}

/** Lista de productos desde la hoja Inventario: [{nombre, precio}]. */
function _productosCatalogo_() {
  var hi = libro().getSheetByName('Inventario');
  var out = [];
  if (hi && hi.getLastRow() >= 2) {
    var ancho = Math.min(8, hi.getMaxColumns());
    hi.getRange(2, 1, hi.getLastRow() - 1, ancho).getValues().forEach(function (f) {
      var nombre = String(f[0] || '').trim();
      if (nombre) out.push({ nombre: nombre, precio: Number(f[7]) || 0 });
    });
  }
  if (!out.length) {
    out = [
      { nombre: 'Agua', precio: PRECIO_AGUA },
      { nombre: 'Cerveza', precio: PRECIO_CERVEZA },
      { nombre: 'Coca Cola', precio: PRECIO_COCA }
    ];
  }
  return out;
}

/**
 * Mapa { idCita: estado } leyendo el Registro en una sola pasada.
 * Igual que idsEnRegistro_() (Codigo.gs) pero trae también la columna 12
 * "Estado", para no tener que releer fila por fila desde el calendario.
 */
function _estadosRegistroPorId_() {
  var hg = libro().getSheetByName('Registro');
  var mapa = {};
  if (!hg || hg.getLastRow() < 2) return mapa;
  var datos = hg.getRange(2, 1, hg.getLastRow() - 1, 16).getValues(); // hasta col 16 = Notas
  datos.forEach(function (f) {
    var nota = String(f[15] || '').trim();
    if (nota.indexOf('Cita ') !== 0) return;
    mapa[nota.substring(5).trim()] = String(f[11] || '').trim(); // col 12 = Estado
  });
  return mapa;
}

/**
 * Busca una cita en la hoja Reservas por su código.
 *
 * @param {Sheet} h - hoja Reservas
 * @param {string} id
 * @returns {Object|null} { fila, f } donde f son las 16 columnas de esa fila
 */
function _filaReserva_(h, id) {
  if (!h || h.getLastRow() < 2) return null;
  id = String(id || '').trim();
  var datos = h.getRange(2, 1, h.getLastRow() - 1, 16).getValues();
  for (var i = 0; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === id) return { fila: i + 2, f: datos[i] };
  }
  return null;
}

/**
 * Cierra un servicio desde el dashboard: escribe Estado, Método de pago,
 * Propina y productos en la fila del Registro, y ajusta el inventario.
 *
 * @param {string} token  - token de sesión (iniciarSesion)
 * @param {string} id     - código de la cita (ej. R260901-AB12C)
 * @param {Object} datos  - { estado, metodo, propina, descuento, servicios:[],
 *                            productos:{Nombre:{v,r}}, desde, hasta }
 *                            `servicios` solo viene si el barbero corrigió lo
 *                            que realmente se hizo; si no, no se toca nada.
 * @returns {Object} { ok, dashboard? } o { ok:false, error }
 */
function cerrarServicioDashboard(token, id, datos) {
  var s;
  try { s = _sesion(token); } catch (e) { return { ok: false, error: e.message }; }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch (e) { return { ok: false, error: 'El sistema está ocupado. Intenta de nuevo.' }; }

  try {
    id = String(id || '').trim();
    var hg = libro().getSheetByName('Registro');
    if (!hg) return { ok: false, error: 'No existe la hoja Registro.' };

    // Asegurar columna 22 "Productos"
    if (hg.getMaxColumns() < 22) hg.insertColumnsAfter(hg.getMaxColumns(), 22 - hg.getMaxColumns());
    if (String(hg.getRange(1, 22).getValue()).trim() !== 'Productos') {
      hg.getRange(1, 22).setValue('Productos')
        .setFontWeight('bold').setBackground('#111111').setFontColor('#FFFFFF');
    }

    var fila = idsEnRegistro_()[id];
    if (!fila) {
      // Todavía no bajó al Registro (cita recién creada, o futura que se quiere
      // cerrar de una vez desde el calendario). volcarAutomatico_() no toma su
      // propio candado -ya estamos dentro del LockService de esta función-,
      // así que se puede llamar aquí sin bloquearse a sí misma.
      try { volcarAutomatico_(); } catch (e) {}
      fila = idsEnRegistro_()[id];
      if (!fila) {
        return { ok: false, error: 'No encontramos esa cita en el Registro.' };
      }
    }

    var barberoFila = String(hg.getRange(fila, 3).getValue()).trim();
    if (s.rol !== 'Dueño' && barberoFila !== s.nombre) {
      return { ok: false, error: 'Esa cita es de ' + barberoFila + ', no la puedes cerrar tú.' };
    }

    datos = datos || {};

    var estado = String(datos.estado || '').trim();
    if (estado) hg.getRange(fila, 12).setValue(estado);

    var metodo = String(datos.metodo || '').trim();
    if (metodo) hg.getRange(fila, 11).setValue(metodo);

    var propina = Number(datos.propina);
    if (isNaN(propina) || propina < 0) propina = 0;
    hg.getRange(fila, 10).setValue(propina);

    // --- Servicios realmente hechos ---
    // Si el barbero los corrigió en el modal, se reescriben aquí ANTES del
    // descuento, porque cambian el valor sobre el que se calcula el Total.
    var hres = libro().getSheetByName(HOJA_RESERVAS);
    var res = _filaReserva_(hres, id);
    var valorServ = Number(hg.getRange(fila, 7).getValue()) || 0;

    if (datos.servicios && datos.servicios.length) {
      var combo = _combinarServicios_(leerConfig(), datos.servicios);
      if (!combo.ok) return { ok: false, error: combo.error };

      hg.getRange(fila, 6).setValue(combo.nombre);
      hg.getRange(fila, 7).setValue(combo.precio);
      hg.getRange(fila, 21).setValue(combo.duracion);
      valorServ = combo.precio;

      // Espejo en Reservas para que la cita no quede inconsistente. NO se tocan
      // fecha ni hora: el servicio ya ocurrió, mover el evento solo confunde.
      if (res) {
        hres.getRange(res.fila, 7).setValue(combo.nombre);
        hres.getRange(res.fila, 8).setValue(combo.duracion);
        hres.getRange(res.fila, 9).setValue(combo.precio);
      }
    }

    // --- Liberar el tiempo sobrante si el servicio terminó antes de lo planeado ---
    // Solo aplica a citas de HOY cerradas como "Atendido". Si el barbero cierra
    // tarde (después de la hora de fin planeada), no se toca nada: la cita nunca
    // se alarga, solo se puede acortar.
    if (estado === 'Atendido' && res) {
      var duracionFinal = (datos.servicios && datos.servicios.length)
        ? combo.duracion
        : (Number(res.f[7]) || 0);           // col8 Reservas = Duración

      if (aISO(res.f[2]) === hoyISO()) {      // col3 Reservas = Fecha
        var iniMin = aMin(aHHMM(res.f[3]));   // col4 Reservas = Hora inicio
        if (iniMin !== null && duracionFinal > 0) {
          var finPlaneadoMin = iniMin + duracionFinal;
          var ahoraMinActual = ahoraMin();
          if (ahoraMinActual > iniMin && ahoraMinActual < finPlaneadoMin) {
            hres.getRange(res.fila, 5).setValue(aTexto(ahoraMinActual)); // col5 = Hora fin
          }
        }
      }
    }

    // Descuento (pesos) — recalcula el Total cobrado (col 9 = valor − descuento)
    var desc = Number(datos.descuento);
    if (isNaN(desc) || desc < 0) desc = 0;
    hg.getRange(fila, 8).setValue(desc);
    if (valorServ > 0) hg.getRange(fila, 9).setValue(Math.max(valorServ - desc, 0));

    // --- Productos: texto en col 22, delta al inventario ---
    // datos.productos = { 'Agua': {v:2, r:0}, 'Gaseosa': {v:1, r:0} }
    var nuevoProd = {};
    var pin = datos.productos || {};
    Object.keys(pin).forEach(function (n) {
      nuevoProd[n] = { v: Number(pin[n] && pin[n].v) || 0, r: Number(pin[n] && pin[n].r) || 0 };
    });
    var viejoProd = _productosDeFila_(hg, fila);

    hg.getRange(fila, 22).setValue(_formatearProductos_(nuevoProd));
    if (hg.getMaxColumns() >= 27) hg.getRange(fila, 23, 1, 5).clearContent();
    _asegurarColValorProductos_(hg);
    hg.getRange(fila, COL_VALOR_PROD).setValue(_valorProductos_(nuevoProd));

    var hi = libro().getSheetByName('Inventario');
    if (hi) actualizarInventario_(hi, _deltaProductos_(nuevoProd, viejoProd));

    // Asegurar las fórmulas de comisión / pago / neto (misma que volcarAutomatico_,
    // en notación A1 — con R1C1 el "Config!C2" de columna entera no se traduce
    // bien y deja la celda en #ERROR!, ver _aplicarFormulasComision_ en Codigo.gs)
    _aplicarFormulasComision_(hg, fila);

    // Dejar marcado en Google Calendar que esta cita ya se cerró.
    if (estado && res) {
      var resumen = String(hg.getRange(fila, 6).getValue()).trim() +
        ' · ' + pesos(Math.max(valorServ - desc, 0)) +
        (metodo ? ' · ' + metodo : '') +
        (propina > 0 ? ' · propina ' + pesos(propina) : '');
      try {
        marcarEventoCerrado_(barberoFila, String(res.f[COL_EVENTO - 1] || '').trim(),
                             estado, resumen);
      } catch (e) {}
    }

    var salida = { ok: true };
    if (datos.desde && datos.hasta) {
      SpreadsheetApp.flush();
      salida.dashboard = getDatosDashboard(
        barberoFila, datos.desde, datos.hasta, s.rol === 'Dueño');
    }
    return salida;

  } catch (err) {
    return { ok: false, error: 'No se pudo guardar el cierre: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/** Trae al Registro las citas que ya terminaron, sin esperar al disparador. */
function traerCitasDeHoy(token) {
  try { _sesion(token); } catch (e) { return { ok: false, error: e.message }; }
  try {
    return { ok: true, n: volcarAutomatico() };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ================================================================
//  FASE C — Crear, modificar y cancelar citas desde el dashboard
// ================================================================

/** Nombres de los barberos REALES: Config + hoja Barberos, menos el calendario general y el/los dueño(s). */
function _barberosReales_() {
  var bs = {};
  try { bs = leerBarberos(); } catch (e) { bs = {}; }
  function esDueno(n) { return bs[n] && _rolDe_(bs[n].rol) === 'Dueño'; }

  var set = {}, lista = [];
  function agregar(n) {
    n = String(n || '').trim();
    if (!n || set[n]) return;
    if (n === CALENDARIO_GENERAL || esDueno(n)) return;
    set[n] = 1; lista.push(n);
  }
  try { leerConfig().barberos.forEach(function (b) { agregar(b.nombre); }); } catch (e) {}
  Object.keys(bs).forEach(agregar);
  return lista;
}

/** Datos para el formulario de "Nueva cita". */
function getFormularioCita(token) {
  var s = _sesion(token);
  var base;
  try { base = getDatosIniciales(); } catch (e) { base = { servicios: [], dias: [] }; }
  return {
    servicios: base.servicios || [],
    barberos:  (s.rol === 'Dueño') ? _barberosReales_() : [s.nombre],
    dias:      base.dias || [],
    rol:       s.rol,
    yo:        s.nombre,
    paso:      PASO_MIN
  };
}

/**
 * Directorio de clientes armado con el historial de Reservas (web y dashboard),
 * para autocompletar la "Nueva cita". Se agrupa por teléfono (últimos 10
 * dígitos); de cada cliente queda el nombre, correo y servicio más recientes.
 * @returns { ok, clientes:[{nombre, telefono, email, visitas, ultima, servicio, barbero}] }
 */
function getClientesDashboard(token) {
  try { _sesion(token); } catch (e) { return { ok: false, error: e.message }; }
  var h = libro().getSheetByName(HOJA_RESERVAS);
  if (!h || h.getLastRow() < 2) return { ok: true, clientes: [] };

  // 0:ID 2:Fecha 5:Barbero 6:Servicio 9:Cliente 10:Teléfono 11:Correo 12:Estado
  var datos = h.getRange(2, 1, h.getLastRow() - 1, 13).getValues();
  var mapa = {};
  datos.forEach(function (f) {
    var nombre = String(f[9] || '').trim();
    var tel = String(f[10] || '').replace(/\D/g, '');
    if (!nombre || nombre === CLIENTE_SIN_NOMBRE || tel.length < 7) return;
    var clave = tel.slice(-10);
    var iso = aISO(f[2]) || '';
    var c = mapa[clave];
    if (!c) c = mapa[clave] = { nombre: '', telefono: tel, email: '', visitas: 0, ultima: '' };
    if (String(f[12] || '').trim() !== 'Cancelada') c.visitas++;
    var email = String(f[11] || '').trim();
    if (email && (!c.email || iso >= c.ultima)) c.email = email;
    if (iso >= c.ultima) {
      c.ultima = iso; c.nombre = nombre; c.telefono = tel;
      c.servicio = String(f[6] || '').trim(); c.barbero = String(f[5] || '').trim();
    }
  });

  var clientes = Object.keys(mapa).map(function (k) { return mapa[k]; });
  clientes.sort(function (a, b) { return a.ultima < b.ultima ? 1 : -1; });
  return { ok: true, clientes: clientes };
}

/** Nombre que queda en la hoja cuando la cita se crea sin cliente (uso interno). */
var CLIENTE_SIN_NOMBRE = 'Sin nombre';

/** "Ahora" redondeado hacia abajo al múltiplo de PASO_MIN. */
function _horaAhoraRedondeada_() {
  return aTexto(Math.floor(ahoraMin() / PASO_MIN) * PASO_MIN);
}

/**
 * ¿La franja barbero+fecha+hora está dentro del horario y sin choque?
 * `excluirId` (opcional): cita que no cuenta como choque (la que se mueve).
 * @returns { ok } o { ok:false, error }
 */
function _validarFranja_(barbero, iso, hhmm, duracion, excluirId) {
  var ini = aMin(hhmm);
  if (ini === null) return { ok: false, error: 'Hora no válida.' };
  var fin = ini + (Number(duracion) || 0);

  var hor = leerHorarios()[barbero];
  var h = hor && hor[diaSemana(iso)];
  if (!h) return { ok: false, error: barbero + ' no trabaja ese día.' };
  var abre = aMin(h.abre), cierra = aMin(h.cierra);
  if (abre !== null && cierra !== null && (ini < abre || fin > cierra)) {
    return { ok: false, error: 'Fuera del horario de ' + barbero + ' (' + h.abre + '–' + h.cierra + ').' };
  }
  var choca = ocupacion(barbero, iso, excluirId).some(function (b) { return ini < b.fin && fin > b.ini; });
  if (choca) return { ok: false, error: barbero + ' ya tiene una cita a esa hora.' };
  return { ok: true };
}

/**
 * Citas próximas (Confirmada) para la pestaña Citas.
 * Barbero: solo las suyas. Dueño: todas.
 */
function getAgendaDashboard(token, dias) {
  var s = _sesion(token);
  dias = Number(dias) || 7;

  var h = libro().getSheetByName(HOJA_RESERVAS);
  if (!h || h.getLastRow() < 2) return { citas: [] };

  var hoy = hoyISO();
  var limite = Utilities.formatDate(new Date(Date.now() + dias * 86400000), TZ, 'yyyy-MM-dd');
  var datos = h.getRange(2, 1, h.getLastRow() - 1, 16).getValues();
  var citas = [];

  datos.forEach(function (f) {
    var id = String(f[0]).trim();
    if (!id) return;
    if (String(f[12]).trim() !== 'Confirmada') return;
    var iso = aISO(f[2]);
    if (!iso || iso < hoy || iso > limite) return;
    var barbero = String(f[5]).trim();
    if (s.rol !== 'Dueño' && barbero !== s.nombre) return;

    citas.push({
      id: id, fecha: iso, hora: aHHMM(f[3]), fin: aHHMM(f[4]),
      barbero: barbero, servicio: String(f[6]).trim(), precio: Number(f[8]) || 0,
      cliente: String(f[9]).trim(), telefono: String(f[10]).trim(),
      notas: String(f[13]).trim(), estado: String(f[12]).trim()
    });
  });

  citas.sort(function (a, b) {
    if (a.fecha !== b.fecha) return a.fecha < b.fecha ? -1 : 1;
    return a.hora < b.hora ? -1 : 1;
  });
  return { citas: citas };
}

/**
 * Datos para la vista de calendario (día/semana/mes) del dashboard.
 * Trae TODAS las citas del rango (incluidas pasadas y canceladas, para poder
 * pintarlas), cruzadas con el estado real del Registro, más horarios,
 * bloqueos y — solo si el rango es corto — eventos manuales de Calendar.
 * Cada hoja se lee una sola vez, sin importar cuántos barberos haya.
 *
 * @param {string} token
 * @param {string} desde - yyyy-mm-dd
 * @param {string} hasta - yyyy-mm-dd
 * @returns {Object} { ok, hoy, barberos, horarios, bloqueos, eventosManuales,
 *                      calendarOmitido, citas[] } o { ok:false, error }
 */
function getCalendarioDashboard(token, desde, hasta) {
  var s;
  try { s = _sesion(token); } catch (e) { return { ok: false, error: e.message }; }

  desde = String(desde || '').slice(0, 10);
  hasta = String(hasta || '').slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(desde) || !/^\d{4}-\d{2}-\d{2}$/.test(hasta) || hasta < desde) {
    return { ok: false, error: 'Rango de fechas no válido.' };
  }

  // Antes de pintar: que Reservas cuadre con el Registro (lo que sale en el
  // Resumen). Si otro proceso tiene el candado, se salta: la rutina de 15 min
  // lo hará igual.
  try {
    var lockSync = LockService.getScriptLock();
    if (lockSync.tryLock(5000)) {
      try { sincronizarReservasConRegistro_(); } finally { lockSync.releaseLock(); }
    }
  } catch (e) { console.error('sync calendario: ' + e); }

  var esDueno = (s.rol === 'Dueño');
  var barberos = esDueno ? _barberosReales_() : [s.nombre];
  var horariosTodos = leerHorarios();

  // --- Citas (una sola lectura de Reservas para todo el rango) ---
  var estados = _estadosRegistroPorId_();
  var citas = [];
  var hr = libro().getSheetByName(HOJA_RESERVAS);
  if (hr && hr.getLastRow() > 1) {
    hr.getRange(2, 1, hr.getLastRow() - 1, 16).getValues().forEach(function (f) {
      var id = String(f[0]).trim();
      if (!id) return;
      var iso = aISO(f[2]);
      if (!iso || iso < desde || iso > hasta) return;
      var barbero = String(f[5]).trim();
      if (!barbero) return;
      if (barberos.indexOf(barbero) < 0) {
        // El dueño ve TODAS las citas: si una está a nombre de alguien que no
        // tiene columna (el mismo dueño, un barbero que ya no está en Config…)
        // se le abre columna, para que nada quede escondido.
        if (!esDueno) return;
        barberos.push(barbero);
      }

      citas.push({
        id: id, fecha: iso, hora: aHHMM(f[3]), fin: aHHMM(f[4]),
        duracion: Number(f[7]) || 0, barbero: barbero,
        servicio: String(f[6]).trim(), precio: Number(f[8]) || 0,
        cliente: String(f[9]).trim(), telefono: String(f[10]).trim(),
        notas: String(f[13]).trim(),
        estadoReserva: String(f[12]).trim(),       // Confirmada | Cancelada
        estadoReal: estados[id] || ''              // Pendiente | Atendido | No asistió | Cancelado | ''
      });
    });
  }
  citas.sort(function (a, b) {
    if (a.fecha !== b.fecha) return a.fecha < b.fecha ? -1 : 1;
    return a.hora < b.hora ? -1 : 1;
  });

  var horarios = {};
  barberos.forEach(function (n) { if (horariosTodos[n]) horarios[n] = horariosTodos[n]; });

  // --- Bloqueos (una sola lectura) ---
  var bloqueos = [];
  var hb = libro().getSheetByName(HOJA_BLOQUEOS);
  if (hb && hb.getLastRow() > 1) {
    hb.getRange(2, 1, hb.getLastRow() - 1, 5).getValues().forEach(function (f) {
      var iso = aISO(f[1]);
      if (!iso || iso < desde || iso > hasta) return;
      var quien = String(f[0]).trim();
      if (quien && quien !== 'Todos' && barberos.indexOf(quien) < 0) return;
      bloqueos.push({
        barbero: quien, fecha: iso,
        desde: aHHMM(f[2]), hasta: aHHMM(f[3]),
        motivo: String(f[4]).trim()
      });
    });
  }

  // --- Eventos manuales de Google Calendar: solo para rangos cortos (día/semana). ---
  // En vista mes se omiten (el mes solo necesita el conteo de citas por día).
  var dias = Math.round((aFechaHora(hasta, '00:00') - aFechaHora(desde, '00:00')) / 86400000) + 1;
  var eventosManuales = {}, calendarOmitido = true;

  if (USAR_CALENDARIO && dias <= 8) {
    calendarOmitido = false;
    var bs = leerBarberos();
    var desdeFecha = aFechaHora(desde, '00:00');
    var hastaFecha = new Date(aFechaHora(hasta, '00:00').getTime() + 86400000);

    barberos.forEach(function (nombre) {
      var b = bs[nombre];
      if (!b || !b.calendario) return;
      eventosManuales[nombre] = [];
      try {
        var cal = CalendarApp.getCalendarById(b.calendario);
        if (!cal) return;
        cal.getEvents(desdeFecha, hastaFecha).forEach(function (ev) {
          if (ev.getTag('riif') === '1') return;   // ya representado en `citas`
          if (ev.isAllDayEvent()) {
            eventosManuales[nombre].push({
              fecha: Utilities.formatDate(ev.getStartTime(), TZ, 'yyyy-MM-dd'),
              ini: '00:00', fin: '23:59', titulo: ev.getTitle()
            });
            return;
          }
          eventosManuales[nombre].push({
            fecha: Utilities.formatDate(ev.getStartTime(), TZ, 'yyyy-MM-dd'),
            ini: Utilities.formatDate(ev.getStartTime(), TZ, 'HH:mm'),
            fin: Utilities.formatDate(ev.getEndTime(), TZ, 'HH:mm'),
            titulo: ev.getTitle()
          });
        });
      } catch (err) { /* si el calendario falla, el resto del calendario sigue funcionando */ }
    });
  }

  return {
    ok: true,
    hoy: hoyISO(),
    barberos: barberos,
    horarios: horarios,
    bloqueos: bloqueos,
    eventosManuales: eventosManuales,
    calendarOmitido: calendarOmitido,
    citas: citas
  };
}

/**
 * Crea una cita desde el dashboard.
 * d = { modo:'ahora'|'agendar', servicios:[nombre,...], barbero, fecha, hora,
 *       cliente, telefono, email, notas }
 */
function crearCitaDashboard(token, d) {
  var s;
  try { s = _sesion(token); } catch (e) { return { ok: false, error: e.message }; }

  d = d || {};
  var cfg = leerConfig();
  var combo = _combinarServicios_(cfg, d.servicios || d.servicio);
  if (!combo.ok) return { ok: false, error: combo.error };

  // Uso interno: nombre y teléfono son opcionales (la reserva web sí los exige).
  var cliente = String(d.cliente || '').trim() || CLIENTE_SIN_NOMBRE;
  var tel = String(d.telefono || '').replace(/\D/g, '');
  if (tel && tel.length < 7) return { ok: false, error: 'El teléfono está incompleto (o déjalo vacío).' };

  // Un barbero solo se agenda a sí mismo; el dueño elige.
  var barbero = (s.rol === 'Dueño') ? String(d.barbero || '').trim() : s.nombre;
  var modo = (d.modo === 'agendar') ? 'agendar' : 'ahora';

  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch (e) { return { ok: false, error: 'El sistema está ocupado. Intenta de nuevo.' }; }

  // Anti-duplicado: el formulario manda una `clave` única por cita. Si llega
  // dos veces (doble toque, o reintento porque el servidor tardó), la segunda
  // devuelve la cita ya creada en vez de crear otra igual.
  var clave = String(d.clave || '').trim();
  var cache = CacheService.getScriptCache();
  if (clave) {
    var previa = cache.get('cita-' + clave);
    if (previa) { lock.releaseLock(); return { ok: true, repetida: true, reserva: { id: previa } }; }
  }

  try {
    var fecha, hora, asignado;

    if (modo === 'agendar') {
      fecha = String(d.fecha || '').slice(0, 10);
      hora  = aHHMM(d.hora);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha) || !hora) {
        return { ok: false, error: 'Falta el día o la hora.' };
      }
      var estado = getCupos(barbero || '*', fecha, combo.duracion);
      var cupo = estado.cupos.filter(function (c) { return c.hora === hora; })[0];
      if (!cupo) return { ok: false, error: 'Ese horario no está libre. Escoge otro.' };
      asignado = (barbero && barbero !== '*' && cupo.barberos.indexOf(barbero) >= 0)
                 ? barbero : cupo.barberos[0];
    } else {
      if (!barbero || barbero === '*') return { ok: false, error: 'Elige el barbero.' };
      fecha = String(d.fecha || '').slice(0, 10);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) fecha = hoyISO();
      hora = aHHMM(d.hora) || _horaAhoraRedondeada_();
      asignado = barbero;

      // El modo "Ahora" también respeta el horario del barbero, salvo forzar.
      if (!d.forzar) {
        var chk = _validarFranja_(barbero, fecha, hora, combo.duracion);
        if (!chk.ok) return { ok: false, fueraHorario: true, error: chk.error };
      }
    }

    var r = {
      id: nuevoIdCita_(),
      fecha: fecha, hora: hora, fin: aTexto(aMin(hora) + combo.duracion),
      barbero: asignado,
      servicio: combo.nombre, precio: combo.precio, duracion: combo.duracion,
      nombre: cliente, telefono: tel, notas: String(d.notas || '').trim()
    };

    // Se marca la clave ANTES de los pasos lentos (Calendar, correos), para
    // que un reintento que llegue mientras tanto tampoco duplique.
    if (clave) cache.put('cita-' + clave, r.id, 21600);   // 6 horas
    return _guardarCitaNueva_(r, { correoCliente: d.email });

  } catch (err) {
    if (clave) { try { cache.remove('cita-' + clave); } catch (e) {} }
    return { ok: false, error: 'No se pudo crear la cita: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/** Lee una cita de Reservas + su estado en el Registro (para el formulario de edición). */
function getCitaDashboard(token, id) {
  var s = _sesion(token);
  id = String(id || '').trim();

  var h = libro().getSheetByName(HOJA_RESERVAS);
  if (!h || h.getLastRow() < 2) return null;
  var datos = h.getRange(2, 1, h.getLastRow() - 1, 16).getValues();

  for (var i = 0; i < datos.length; i++) {
    if (String(datos[i][0]).trim() !== id) continue;
    var f = datos[i];
    var barbero = String(f[5]).trim();
    if (s.rol !== 'Dueño' && barbero !== s.nombre) return { error: 'Esa cita no es tuya.' };

    var filaReg = idsEnRegistro_()[id];
    var estadoReg = '', metodoReg = '', propinaReg = 0, descuentoReg = 0, productosReg = {};
    if (filaReg) {
      var hg = libro().getSheetByName('Registro');
      estadoReg    = String(hg.getRange(filaReg, 12).getValue()).trim();
      metodoReg    = String(hg.getRange(filaReg, 11).getValue()).trim();
      propinaReg   = Number(hg.getRange(filaReg, 10).getValue()) || 0;
      descuentoReg = Number(hg.getRange(filaReg, 8).getValue()) || 0;
      productosReg = _productosDeFila_(hg, filaReg);
    }
    return {
      id: id, fecha: aISO(f[2]), hora: aHHMM(f[3]), fin: aHHMM(f[4]),
      barbero: barbero, servicio: String(f[6]).trim(),
      duracion: Number(f[7]) || 0, precio: Number(f[8]) || 0,
      cliente: String(f[9]).trim(), telefono: String(f[10]).trim(),
      email: String(f[11]).trim(), notas: String(f[13]).trim(),
      estado: String(f[12]).trim(),
      enRegistro: !!filaReg, estadoRegistro: estadoReg,
      metodoRegistro: metodoReg, propinaRegistro: propinaReg,
      descuentoRegistro: descuentoReg, productosRegistro: productosReg
    };
  }
  return null;
}

/**
 * Modifica una cita. `cambios` trae solo los campos que cambian:
 * { fecha, hora, barbero, servicio, cliente, telefono, notas }
 * `cambios.mover` = viene de arrastrar la cita en el calendario: se valida
 * que la nueva franja esté en el horario y libre, y si el que la mueve es el
 * mismo barbero no se le manda aviso (ya lo sabe).
 */
function modificarCitaDashboard(token, id, cambios) {
  var s;
  try { s = _sesion(token); } catch (e) { return { ok: false, error: e.message }; }

  id = String(id || '').trim();
  cambios = cambios || {};

  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch (e) { return { ok: false, error: 'El sistema está ocupado. Intenta de nuevo.' }; }

  try {
    var h = libro().getSheetByName(HOJA_RESERVAS);
    if (!h || h.getLastRow() < 2) return { ok: false, error: 'No encontramos esa cita.' };
    var datos = h.getRange(2, 1, h.getLastRow() - 1, 16).getValues();

    for (var i = 0; i < datos.length; i++) {
      if (String(datos[i][0]).trim() !== id) continue;
      var f = datos[i];
      var fila = i + 2;

      var barberoViejo = String(f[5]).trim();
      if (s.rol !== 'Dueño' && barberoViejo !== s.nombre) {
        return { ok: false, error: 'Esa cita es de ' + barberoViejo + '.' };
      }
      if (String(f[12]).trim() === 'Cancelada') {
        return { ok: false, error: 'Esa cita ya está cancelada.' };
      }

      var hg = libro().getSheetByName('Registro');
      var filaReg = idsEnRegistro_()[id];
      var estadoReg = filaReg ? String(hg.getRange(filaReg, 12).getValue()).trim() : '';
      if (estadoReg === 'Atendido') {
        return { ok: false, error: 'Esa cita ya se cerró. Para corregir el valor usa la pantalla de cierre.' };
      }

      var cfg = leerConfig();
      var combo;
      if (cambios.servicios && cambios.servicios.length) {
        // El barbero eligió servicios nuevos: deben existir en Config.
        combo = _combinarServicios_(cfg, cambios.servicios);
        if (!combo.ok) return { ok: false, error: combo.error };
      } else {
        // No tocó los servicios: tratar de reconstruir el combo actual para
        // no perder precio/duración si viene de una cita con varios servicios.
        combo = _combinarServicios_(cfg, String(f[6]).trim().split(SEPARADOR_COMBO));
        if (!combo.ok) {
          // Nombre viejo o escrito a mano que ya no coincide con Config:
          // se conserva tal cual en vez de inventar un valor.
          combo = { ok: true, nombre: String(f[6]).trim(), precio: Number(f[8]) || 0, duracion: Number(f[7]) || 0 };
        }
      }

      var fecha   = cambios.fecha ? String(cambios.fecha).slice(0, 10) : aISO(f[2]);
      var hora    = cambios.hora ? aHHMM(cambios.hora) : aHHMM(f[3]);
      var barbero = cambios.barbero ? String(cambios.barbero).trim() : barberoViejo;
      var cliente = (cambios.cliente != null && cambios.cliente !== '')
                    ? String(cambios.cliente).trim() : String(f[9]).trim();
      var tel     = (cambios.telefono != null && cambios.telefono !== '')
                    ? String(cambios.telefono).replace(/\D/g, '') : String(f[10]).trim();
      var notas   = (cambios.notas != null) ? String(cambios.notas).trim() : String(f[13]).trim();

      if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha) || !hora) {
        return { ok: false, error: 'Día u hora no válidos.' };
      }
      if (!cliente) cliente = CLIENTE_SIN_NOMBRE;

      if (cambios.mover) {
        var chk = _validarFranja_(barbero, fecha, hora, combo.duracion || PASO_MIN, id);
        if (!chk.ok) return { ok: false, error: chk.error };
      }

      var finTxt = aTexto(aMin(hora) + combo.duracion);
      var r = {
        id: id, fecha: fecha, hora: hora, fin: finTxt, barbero: barbero,
        servicio: combo.nombre, precio: combo.precio, duracion: combo.duracion,
        nombre: cliente, telefono: tel, notas: notas
      };

      // Reservas
      h.getRange(fila, 3).setValue(fecha);
      h.getRange(fila, 4).setValue(hora);
      h.getRange(fila, 5).setValue(finTxt);
      h.getRange(fila, 6).setValue(barbero);
      h.getRange(fila, 7).setValue(combo.nombre);
      h.getRange(fila, 8).setValue(combo.duracion);
      h.getRange(fila, 9).setValue(combo.precio);
      h.getRange(fila, 10).setValue(cliente);
      h.getRange(fila, 11).setValue(tel);
      h.getRange(fila, 14).setValue(notas);

      // Calendario: borrar el viejo, crear el nuevo
      try { borrarEvento(barberoViejo, String(f[14] || '').trim()); } catch (e) {}
      var idEvento = '';
      try { idEvento = crearEvento(r); } catch (e) {}
      h.getRange(fila, COL_EVENTO).setValue(idEvento || '');

      var link = linkAvisoBarbero(r);
      h.getRange(fila, COL_AVISO).setFormula(
        link ? '=HYPERLINK("' + link + '";"Avisar a ' + r.barbero + '")' : '');

      var avisar = !(cambios.mover && s.nombre === barbero && barberoViejo === barbero);
      if (avisar) { try { notificarBarbero(r, 'modificada'); } catch (e) {} }

      // Registro (si está y no Atendido)
      if (filaReg) {
        var p = fecha.split('-');
        var fechaReal = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
        var descAct = Number(hg.getRange(filaReg, 8).getValue()) || 0;
        hg.getRange(filaReg, 1).setValue(fechaReal);
        hg.getRange(filaReg, 1).setNumberFormat('DD/MM/YYYY');
        hg.getRange(filaReg, 2).setValue(hora);
        hg.getRange(filaReg, 3).setValue(barbero);
        hg.getRange(filaReg, 4).setValue(cliente);
        hg.getRange(filaReg, 5).setValue(tel);
        hg.getRange(filaReg, 6).setValue(combo.nombre);
        hg.getRange(filaReg, 7).setValue(combo.precio);
        hg.getRange(filaReg, 9).setValue(Math.max(combo.precio - descAct, 0));
        hg.getRange(filaReg, 17).setValue(fechaReal.getFullYear());
        hg.getRange(filaReg, 18).setValue(fechaReal.getMonth() + 1);
        hg.getRange(filaReg, 19).setValue(DIAS[(fechaReal.getDay() + 6) % 7]);
        hg.getRange(filaReg, 20).setValue(parseInt(hora.split(':')[0], 10));
        hg.getRange(filaReg, 21).setValue(combo.duracion);
      }

      return { ok: true };
    }
    return { ok: false, error: 'No encontramos esa cita.' };

  } catch (err) {
    return { ok: false, error: 'No se pudo modificar: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/** Cancela una cita desde el dashboard (verifica que sea del barbero o que sea el dueño). */
function cancelarCitaDashboard(token, id) {
  var s;
  try { s = _sesion(token); } catch (e) { return { ok: false, error: e.message }; }
  id = String(id || '').trim();

  var c = getCitaDashboard(token, id);
  if (!c) return { ok: false, error: 'No encontramos esa cita.' };
  if (c.error) return { ok: false, error: c.error };
  if (s.rol !== 'Dueño' && c.barbero !== s.nombre) return { ok: false, error: 'Esa cita no es tuya.' };
  // Una cita ya cobrada no se "cancela": el Registro la seguiría contando como
  // Atendido (plata y comisión) aunque en Reservas y el calendario desaparezca.
  if (c.estadoRegistro === 'Atendido') {
    return { ok: false, error: s.rol === 'Dueño'
      ? 'Esta cita ya se cobró. Si fue un error o está duplicada, usa "Borrar definitivamente".'
      : 'Esta cita ya se cobró. Pídele al administrador que la elimine si fue un error.' };
  }
  return cancelarReserva(id);
}

/**
 * Elimina una cita de TODO el sistema (solo Dueño): fila de Reservas, fila del
 * Registro, evento de Google Calendar, y devuelve al inventario los productos
 * que se habían descontado. Pensado para citas duplicadas o creadas por error;
 * para un cliente que no vino se usa "Cancelar", que deja la huella.
 */
function eliminarCitaDashboard(token, id) {
  try { _sesionDueno_(token); } catch (e) { return { ok: false, error: e.message }; }
  id = String(id || '').trim();
  if (!id) return { ok: false, error: 'Falta el código de la cita.' };

  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch (e) { return { ok: false, error: 'El sistema está ocupado. Intenta de nuevo.' }; }

  try {
    var encontrada = false;

    // 1. Registro: devolver productos al inventario y borrar la fila
    var hg = libro().getSheetByName('Registro');
    var filaReg = idsEnRegistro_()[id];
    if (hg && filaReg) {
      var prod = _productosDeFila_(hg, filaReg);
      var hi = libro().getSheetByName('Inventario');
      if (hi) actualizarInventario_(hi, _deltaProductos_({}, prod));
      hg.deleteRow(filaReg);
      encontrada = true;
    }

    // 2. Reservas + evento de Calendar
    var h = libro().getSheetByName(HOJA_RESERVAS);
    if (h && h.getLastRow() >= 2) {
      var datos = h.getRange(2, 1, h.getLastRow() - 1, COL_EVENTO).getValues();
      for (var i = datos.length - 1; i >= 0; i--) {
        if (String(datos[i][0]).trim() !== id) continue;
        var idEvento = String(datos[i][COL_EVENTO - 1] || '').trim();
        if (idEvento) { try { borrarEvento(String(datos[i][5]).trim(), idEvento); } catch (e) {} }
        h.deleteRow(i + 2);
        encontrada = true;
      }
    }

    if (!encontrada) return { ok: false, error: 'No encontramos esa cita.' };
    SpreadsheetApp.flush();
    return { ok: true };
  } catch (err) {
    return { ok: false, error: 'No se pudo eliminar: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/** Link de WhatsApp para avisarle al barbero de una cita (botón "Avisar al barbero"). */
function notificarCitaDashboard(token, id) {
  try { _sesion(token); } catch (e) { return { ok: false, error: e.message }; }
  var c = getCitaDashboard(token, id);
  if (!c || c.error) return { ok: false, error: (c && c.error) || 'No encontramos esa cita.' };
  var link = linkAvisoBarbero({
    id: c.id, barbero: c.barbero, servicio: c.servicio, fecha: c.fecha,
    hora: c.hora, fin: c.fin, duracion: c.duracion, precio: c.precio,
    nombre: c.cliente, telefono: c.telefono, notas: c.notas
  });
  if (!link) return { ok: false, error: 'Ese barbero no tiene WhatsApp cargado en la hoja Barberos.' };
  return { ok: true, link: link };
}

// ================================================================
//  FASE D — Inventario y productos (solo Dueño)
// ================================================================

/** Valida sesión y exige rol Dueño. Devuelve la sesión o lanza. */
function _sesionDueno_(token) {
  var s = _sesion(token);
  if (s.rol !== 'Dueño') throw new Error('Solo el administrador puede ver esto.');
  return s;
}

/**
 * Inventario completo para la pestaña del dueño.
 * Stock = inicial + entradas − vendidas − cortesías + ajustes (col I).
 * Si en la hoja alguien escribió un número a mano en "Stock actual" (col F,
 * borrando la fórmula), ese número se respeta: se pasa la diferencia a
 * Ajustes y se vuelve a poner la fórmula, para que siga descontando ventas.
 */
function getInventarioDashboard(token) {
  _sesionDueno_(token);
  var hi = libro().getSheetByName('Inventario');
  if (!hi || hi.getLastRow() < 2) return { productos: [] };
  _asegurarColAjustes_(hi);

  var n = hi.getLastRow() - 1;
  var filas = hi.getRange(2, 1, n, 9).getValues();
  var formulasF = hi.getRange(2, 6, n, 1).getFormulas();
  var productos = [];
  filas.forEach(function (f, i) {
    var nombre = String(f[0] || '').trim();
    if (!nombre) return;
    var ini = Number(f[1]) || 0, ent = Number(f[2]) || 0;
    var ven = Number(f[3]) || 0, reg = Number(f[4]) || 0;
    var aju = Number(f[8]) || 0;
    var stock = ini + ent - ven - reg + aju;

    var fila = i + 2;
    if (!formulasF[i][0]) {
      var manual = f[5];
      if (manual !== '' && manual !== null && !isNaN(Number(manual)) && Number(manual) !== stock) {
        aju += Number(manual) - stock;
        stock = Number(manual);
        hi.getRange(fila, 9).setValue(aju);
      }
      hi.getRange(fila, 6).setFormula(_formulaStock_(fila));
    }

    productos.push({
      producto: nombre,
      precio: Number(f[7]) || 0,
      stockActual: stock,
      vendidas: ven,
      regaladas: reg,
      valorVendido: Number(f[6]) || 0,
      bajo: stock <= 3
    });
  });
  return { productos: productos };
}

/**
 * Deja el stock de un producto en `cantidad` exacta (conteo físico). La
 * diferencia va a la col I "Ajustes", así no se pierde el historial de
 * entradas, ventas ni cortesías. Solo Dueño.
 */
function ajustarStockDashboard(token, producto, cantidad) {
  try { _sesionDueno_(token); } catch (e) { return { ok: false, error: e.message }; }
  cantidad = Number(cantidad);
  if (isNaN(cantidad) || cantidad < 0 || Math.floor(cantidad) !== cantidad) {
    return { ok: false, error: 'Escribe una cantidad entera (0 o más).' };
  }
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (e) { return { ok: false, error: 'Intenta de nuevo.' }; }
  try {
    var hi = libro().getSheetByName('Inventario');
    if (!hi || hi.getLastRow() < 2) return { ok: false, error: 'No hay inventario.' };
    _asegurarColAjustes_(hi);
    var filas = hi.getRange(2, 1, hi.getLastRow() - 1, 9).getValues();
    for (var i = 0; i < filas.length; i++) {
      var f = filas[i];
      if (normalizar(f[0]) !== normalizar(producto)) continue;
      var fila = i + 2;
      var sinAjuste = (Number(f[1]) || 0) + (Number(f[2]) || 0) - (Number(f[3]) || 0) - (Number(f[4]) || 0);
      hi.getRange(fila, 9).setValue(cantidad - sinAjuste);
      hi.getRange(fila, 6).setFormula(_formulaStock_(fila));
      SpreadsheetApp.flush();
      return { ok: true, inventario: getInventarioDashboard(token) };
    }
    return { ok: false, error: 'No encontré ese producto.' };
  } finally {
    lock.releaseLock();
  }
}

/** Repone stock (suma a Entradas). Solo Dueño. */
function reponerStockDashboard(token, producto, cantidad) {
  try { _sesionDueno_(token); } catch (e) { return { ok: false, error: e.message }; }
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (e) { return { ok: false, error: 'Intenta de nuevo.' }; }
  try {
    var r = _reponerStock_(producto, cantidad);
    if (!r.ok) return r;
    SpreadsheetApp.flush();
    return { ok: true, inventario: getInventarioDashboard(token) };
  } finally {
    lock.releaseLock();
  }
}

/** Crea un producto nuevo en la hoja Inventario. Solo Dueño. */
function crearProductoDashboard(token, datos) {
  try { _sesionDueno_(token); } catch (e) { return { ok: false, error: e.message }; }
  datos = datos || {};
  var nombre = String(datos.nombre || '').replace(/[:;]/g, ' ').trim();
  var precio = Number(datos.precio);
  var stockIni = Number(datos.stockInicial) || 0;
  if (nombre.length < 2) return { ok: false, error: 'Escribe el nombre del producto.' };
  if (isNaN(precio) || precio < 0) return { ok: false, error: 'Precio inválido.' };

  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (e) { return { ok: false, error: 'Intenta de nuevo.' }; }
  try {
    var hi = libro().getSheetByName('Inventario');
    if (!hi) return { ok: false, error: 'No existe la hoja Inventario. Corre "Configurar inventario" desde el menú.' };

    // ¿ya existe?
    var filas = hi.getLastRow() >= 2 ? hi.getRange(2, 1, hi.getLastRow() - 1, 1).getValues() : [];
    for (var i = 0; i < filas.length; i++) {
      if (normalizar(filas[i][0]) === normalizar(nombre)) {
        return { ok: false, error: 'Ya existe un producto "' + String(filas[i][0]).trim() + '".' };
      }
    }

    var fila = hi.getLastRow() + 1;
    hi.getRange(fila, 1, 1, 8).setValues([[nombre, stockIni, 0, 0, 0, '', 0, precio]]);
    _asegurarColAjustes_(hi);
    hi.getRange(fila, 6).setFormula(_formulaStock_(fila));
    hi.getRange(fila, 7).setNumberFormat('$#,##0');
    hi.getRange(fila, 8).setNumberFormat('$#,##0');
    SpreadsheetApp.flush();
    return { ok: true, inventario: getInventarioDashboard(token) };
  } finally {
    lock.releaseLock();
  }
}

/** Cambia el precio de un producto. Solo Dueño. */
function setPrecioProductoDashboard(token, producto, precio) {
  try { _sesionDueno_(token); } catch (e) { return { ok: false, error: e.message }; }
  precio = Number(precio);
  if (isNaN(precio) || precio < 0) return { ok: false, error: 'Precio inválido.' };

  var hi = libro().getSheetByName('Inventario');
  if (!hi || hi.getLastRow() < 2) return { ok: false, error: 'No hay inventario.' };
  var filas = hi.getRange(2, 1, hi.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < filas.length; i++) {
    if (normalizar(filas[i][0]) === normalizar(producto)) {
      hi.getRange(i + 2, 8).setValue(precio);
      SpreadsheetApp.flush();
      return { ok: true, inventario: getInventarioDashboard(token) };
    }
  }
  return { ok: false, error: 'No encontré ese producto.' };
}
