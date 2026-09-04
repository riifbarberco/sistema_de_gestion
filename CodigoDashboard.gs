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

  // Leer todo el Registro. Se piden hasta 27 columnas (productos 22-27) pero se
  // respeta el ancho real de la hoja para no salirse de la cuadrícula.
  var anchoLeer = Math.max(16, Math.min(27, hg.getLastColumn()));
  var datos = hg.getRange(2, 1, hg.getLastRow() - 1, anchoLeer).getValues();

  var servicios     = [];
  var totalCobrado  = 0;
  var totalComision = 0;
  var totalPropinas = 0;
  var contPorDia    = {};   // { 'YYYY-MM-DD': { count, total } }
  var contServicios = {};   // { 'Nombre servicio': { count, total } }

  datos.forEach(function(f) {
    // Columnas del Registro (base-1 en Sheets → base-0 aquí):
    //  0:Fecha  1:Hora  2:Barbero  3:Cliente  4:Teléfono  5:Servicio
    //  6:Valor  7:Desc  8:Total    9:Propina  10:Método   11:Estado
    // 12:Comisión  13:PagoBarbero  14:Neto  15:Notas
    // 16:Año  17:Mes  18:DiaSem  19:HoraBloque  20:Duración

    var fBarbero = String(f[2] || '').trim();
    if (!verTodo && fBarbero !== barbero) return;

    var estado = String(f[11] || '').trim();
    if (estado === 'Cancelado') return; // ignorar canceladas

    // Obtener fecha ISO de la columna Fecha (puede ser Date o string)
    var iso = '';
    if (f[0] instanceof Date) {
      iso = aISO(f[0]);
    } else {
      var s = String(f[0] || '').trim();
      // acepta DD/MM/YYYY y YYYY-MM-DD
      var mDMY = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
      var mYMD = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
      if (mDMY) {
        iso = mDMY[3] + '-' + ('0'+mDMY[2]).slice(-2) + '-' + ('0'+mDMY[1]).slice(-2);
      } else if (mYMD) {
        iso = s.slice(0, 10);
      } else {
        return; // fecha inválida, saltar fila
      }
    }

    if (iso < desde || iso > hasta) return;

    var total       = Number(f[8])  || 0;
    var descuento   = Number(f[7])  || 0;
    var propina     = Number(f[9])  || 0;
    var comision    = Number(f[12]) || 0;
    var pagoBarbero = Number(f[13]) || 0;
    var nombreSrv   = String(f[5] || '').trim();
    var hora        = String(f[1] || '').trim();
    var cliente     = String(f[3] || '').trim();
    var metodo      = String(f[10] || '').trim();

    // id de la cita, guardado en la col 16 (Notas) como "Cita <id>"
    var nota = String(f[15] || '').trim();
    var citaId = (nota.indexOf('Cita ') === 0) ? nota.substring(5).trim() : '';

    // Productos: { 'Agua': {v,r}, ... } — acepta formato nuevo (texto) y viejo.
    var productos = _leerProductosRegistro_([f[21], f[22], f[23], f[24], f[25], f[26]]);

    totalCobrado  += total;
    totalComision += pagoBarbero;   // "pago barbero" = comisión + propina
    totalPropinas += propina;

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
      total: total, pagoBarbero: pagoBarbero,
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
    totalComision:  Math.round(totalComision),
    totalPropinas:  Math.round(totalPropinas),
    porPeriodo:     porPeriodo,
    topServicios:   topServicios,
    servicios:      servicios
  };
}

/** Resultado vacío cuando no hay datos */
function _dashboardVacio() {
  return {
    totalServicios: 0, totalCobrado: 0,
    totalComision: 0,  totalPropinas: 0,
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

  if (!metodos.length) metodos = ['Efectivo', 'Nequi', 'Daviplata', 'Transferencia', 'Tarjeta'];
  if (!estados.length) estados = ['Atendido', 'No asistió', 'Cancelado'];

  return { metodos: metodos, estados: estados, productos: productos };
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
 * Cierra un servicio desde el dashboard: escribe Estado, Método de pago,
 * Propina y productos en la fila del Registro, y ajusta el inventario.
 *
 * @param {string} token  - token de sesión (iniciarSesion)
 * @param {string} id     - código de la cita (ej. R260901-AB12C)
 * @param {Object} datos  - { estado, metodo, propina, productos:{aguaV,aguaR,
 *                            cervV,cervR,cocaV,cocaR}, desde, hasta }
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
      return { ok: false, error: 'Esa cita todavía no está en el Registro. Toca "Actualizar" y vuelve a intentar.' };
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
    hg.getRange(fila, 10).setValue(isNaN(propina) || propina < 0 ? 0 : propina);

    // Descuento (pesos) — recalcula el Total cobrado (col 9 = valor − descuento)
    var desc = Number(datos.descuento);
    if (isNaN(desc) || desc < 0) desc = 0;
    var valorServ = Number(hg.getRange(fila, 7).getValue()) || 0;
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

    var hi = libro().getSheetByName('Inventario');
    if (hi) actualizarInventario_(hi, _deltaProductos_(nuevoProd, viejoProd));

    // Asegurar las fórmulas de comisión / pago / neto (idénticas a volcarAutomatico)
    hg.getRange(fila, 13).setFormulaR1C1(
      '=IF(RC12="Atendido",IFERROR(RC9*INDEX(Config!C2,MATCH(RC3,Config!C1,0)),0),0)');
    hg.getRange(fila, 14).setFormulaR1C1(
      '=IF(RC12="Atendido",RC13+IF(RC10="",0,RC10),0)');
    hg.getRange(fila, 15).setFormulaR1C1(
      '=IF(RC12="Atendido",RC9-RC13,0)');

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

/** "Ahora" redondeado hacia abajo al múltiplo de PASO_MIN. */
function _horaAhoraRedondeada_() {
  return aTexto(Math.floor(ahoraMin() / PASO_MIN) * PASO_MIN);
}

/**
 * ¿La franja barbero+fecha+hora está dentro del horario y sin choque?
 * @returns { ok } o { ok:false, error }
 */
function _validarFranja_(barbero, iso, hhmm, duracion) {
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
  var choca = ocupacion(barbero, iso).some(function (b) { return ini < b.fin && fin > b.ini; });
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
 * Crea una cita desde el dashboard.
 * d = { modo:'ahora'|'agendar', servicio, barbero, fecha, hora, cliente,
 *       telefono, email, notas }
 */
function crearCitaDashboard(token, d) {
  var s;
  try { s = _sesion(token); } catch (e) { return { ok: false, error: e.message }; }

  d = d || {};
  var cfg = leerConfig();
  var servicio = cfg.servicios.filter(function (x) { return x.nombre === d.servicio; })[0];
  if (!servicio) return { ok: false, error: 'Ese servicio ya no está disponible.' };

  var cliente = String(d.cliente || '').trim();
  var tel = String(d.telefono || '').replace(/\D/g, '');
  if (cliente.length < 3) return { ok: false, error: 'Escribe el nombre del cliente.' };
  if (tel.length < 7) return { ok: false, error: 'Escribe un teléfono válido.' };

  // Un barbero solo se agenda a sí mismo; el dueño elige.
  var barbero = (s.rol === 'Dueño') ? String(d.barbero || '').trim() : s.nombre;
  var modo = (d.modo === 'agendar') ? 'agendar' : 'ahora';

  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch (e) { return { ok: false, error: 'El sistema está ocupado. Intenta de nuevo.' }; }

  try {
    var fecha, hora, asignado;

    if (modo === 'agendar') {
      fecha = String(d.fecha || '').slice(0, 10);
      hora  = aHHMM(d.hora);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha) || !hora) {
        return { ok: false, error: 'Falta el día o la hora.' };
      }
      var estado = getCupos(barbero || '*', fecha, servicio.duracion);
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
        var chk = _validarFranja_(barbero, fecha, hora, servicio.duracion);
        if (!chk.ok) return { ok: false, fueraHorario: true, error: chk.error };
      }
    }

    var r = {
      id: nuevoIdCita_(),
      fecha: fecha, hora: hora, fin: aTexto(aMin(hora) + servicio.duracion),
      barbero: asignado,
      servicio: servicio.nombre, precio: servicio.precio, duracion: servicio.duracion,
      nombre: cliente, telefono: tel, notas: String(d.notas || '').trim()
    };

    return _guardarCitaNueva_(r, { correoCliente: d.email });

  } catch (err) {
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
    var estadoReg = '';
    if (filaReg) {
      estadoReg = String(libro().getSheetByName('Registro').getRange(filaReg, 12).getValue()).trim();
    }
    return {
      id: id, fecha: aISO(f[2]), hora: aHHMM(f[3]), fin: aHHMM(f[4]),
      barbero: barbero, servicio: String(f[6]).trim(),
      duracion: Number(f[7]) || 0, precio: Number(f[8]) || 0,
      cliente: String(f[9]).trim(), telefono: String(f[10]).trim(),
      email: String(f[11]).trim(), notas: String(f[13]).trim(),
      estado: String(f[12]).trim(),
      enRegistro: !!filaReg, estadoRegistro: estadoReg
    };
  }
  return null;
}

/**
 * Modifica una cita. `cambios` trae solo los campos que cambian:
 * { fecha, hora, barbero, servicio, cliente, telefono, notas }
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
      var servNombre = cambios.servicio ? String(cambios.servicio).trim() : String(f[6]).trim();
      var serv = cfg.servicios.filter(function (x) { return x.nombre === servNombre; })[0];
      if (!serv) return { ok: false, error: 'Ese servicio no existe.' };

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
      if (cliente.length < 3) return { ok: false, error: 'El nombre del cliente es muy corto.' };

      var finTxt = aTexto(aMin(hora) + serv.duracion);
      var r = {
        id: id, fecha: fecha, hora: hora, fin: finTxt, barbero: barbero,
        servicio: serv.nombre, precio: serv.precio, duracion: serv.duracion,
        nombre: cliente, telefono: tel, notas: notas
      };

      // Reservas
      h.getRange(fila, 3).setValue(fecha);
      h.getRange(fila, 4).setValue(hora);
      h.getRange(fila, 5).setValue(finTxt);
      h.getRange(fila, 6).setValue(barbero);
      h.getRange(fila, 7).setValue(serv.nombre);
      h.getRange(fila, 8).setValue(serv.duracion);
      h.getRange(fila, 9).setValue(serv.precio);
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
        link ? '=HYPERLINK("' + link + '","Avisar a ' + r.barbero + '")' : '');

      try { notificarBarbero(r, 'modificada'); } catch (e) {}

      // Registro (si está y no Atendido)
      if (filaReg) {
        var p = fecha.split('-');
        var fechaReal = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
        var descAct = Number(hg.getRange(filaReg, 8).getValue()) || 0;
        hg.getRange(filaReg, 1).setValue(fechaReal);
        hg.getRange(filaReg, 1).setNumberFormat('DD/MM/YYYY');
        hg.getRange(filaReg, 2).setValue(hora);
        hg.getRange(filaReg, 3).setValue(barbero);
        hg.getRange(filaReg, 6).setValue(serv.nombre);
        hg.getRange(filaReg, 7).setValue(serv.precio);
        hg.getRange(filaReg, 9).setValue(Math.max(serv.precio - descAct, 0));
        hg.getRange(filaReg, 17).setValue(fechaReal.getFullYear());
        hg.getRange(filaReg, 18).setValue(fechaReal.getMonth() + 1);
        hg.getRange(filaReg, 19).setValue(DIAS[(fechaReal.getDay() + 6) % 7]);
        hg.getRange(filaReg, 20).setValue(parseInt(hora.split(':')[0], 10));
        hg.getRange(filaReg, 21).setValue(serv.duracion);
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

  if (s.rol !== 'Dueño') {
    var c = getCitaDashboard(token, id);
    if (!c) return { ok: false, error: 'No encontramos esa cita.' };
    if (c.error) return { ok: false, error: c.error };
    if (c.barbero !== s.nombre) return { ok: false, error: 'Esa cita no es tuya.' };
  }
  return cancelarReserva(id);
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

/** Inventario completo para la pestaña del dueño. */
function getInventarioDashboard(token) {
  _sesionDueno_(token);
  var hi = libro().getSheetByName('Inventario');
  if (!hi || hi.getLastRow() < 2) return { productos: [] };

  var filas = hi.getRange(2, 1, hi.getLastRow() - 1, 8).getValues();
  var productos = [];
  filas.forEach(function (f) {
    var nombre = String(f[0] || '').trim();
    if (!nombre) return;
    var ini = Number(f[1]) || 0, ent = Number(f[2]) || 0;
    var ven = Number(f[3]) || 0, reg = Number(f[4]) || 0;
    var stock = ini + ent - ven - reg;
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
    hi.getRange(fila, 6).setFormulaR1C1('=RC2+RC3-RC4-RC5');
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
