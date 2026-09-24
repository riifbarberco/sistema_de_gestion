/**
 * RIIF — Sistema de reservas propio
 * v2 — Google Calendar + notificaciones a barberos
 *
 * NUEVO EN ESTA VERSIÓN
 *  · Cada reserva crea un evento en el calendario del barbero
 *  · El barbero queda invitado: le llega a SU Google Calendar y a su correo
 *  · Correo de aviso al barbero con todos los datos de la cita
 *  · WhatsApp opcional al barbero (ver la guía antes de activarlo)
 *  · Bidireccional: lo que bloquees en Google Calendar desaparece del formulario
 *  · Al cancelar, el evento se borra y el barbero se entera
 *
 * ANTES DE PEGAR ESTE ARCHIVO: copia los valores que ya tenías en tu bloque
 * AJUSTES y vuelve a ponerlos abajo. Después corre la función actualizarSistema.
 *
 * IMPORTANTE: en Configuración del proyecto, verifica que la zona horaria
 * del script sea (GMT-05:00) Bogotá.
 */
 

// ==================== AJUSTES ====================
var NEGOCIO        = 'RIIF';
var TZ             = 'America/Bogota';
var WHATSAPP       = '573015857963';   // tu WhatsApp: 57 + número, sin + ni espacios
var EMAIL_AVISO    = 'riifbarberco@gmail.com';  // tu correo, recibe aviso de cada cita. Vacío = no envía
var DIRECCION      = 'Riif Barber Co. - Cartagena';
var PASO_MIN       = 30;
var ANTICIPACION_H = 1;
var DIAS_ADELANTE  = 21;


// --- Calendario ---
var USAR_CALENDARIO    = true;   // false apaga todo lo de Google Calendar
var PREFIJO_CALENDARIO = 'RIIF · ';
// Fila de la hoja Barberos que NO es un barbero sino el calendario donde se
// junta todo. Deja '' si no quieres calendario general.
var CALENDARIO_GENERAL = 'Riif Reservas';
var RECORDATORIO_MIN   = 60;     // aviso al barbero, minutos antes de la cita
 
// --- Qué ve el barbero ---
// false = el barbero NO recibe el teléfono del cliente en su calendario ni en
// su correo. Úsalo si quieres que la relación con el cliente sea del negocio.
var MOSTRAR_TELEFONO_A_BARBERO = false;
 
// --- Formulario de cierre de servicio (para los barberos) ---
var FORM_CIERRE_URL = 'https://docs.google.com/forms/d/e/1FAIpQLSfO5550qY3ES8xbE1TdrxZKNR9ANjFh0feXCh0TlbqpVXgkCQ/viewform';
var FORM_ENTRY_CODIGO = 'entry.984619231';   // campo Código de cita
var FORM_ENTRY_ESTADO = 'entry.894483872';   // campo Estado
var FORM_ENTRY_METODO = 'entry.1058757202';  // campo Método de pago
// ID de la hoja de respuestas del form (para leer y actualizar el Registro)
var FORM_SHEET_ID = '16TTpvhxXIVhJMkqyu2lvs9bKKgnHsAms1yigGx5vC08';
 
// --- Botón de WhatsApp de la pantalla de confirmación ---
// 'negocio' = el mensaje llega al WhatsApp del local (recomendado)
// 'barbero' = el mensaje llega directo al celular del barbero asignado
var WHATSAPP_DESTINO = 'negocio';
 
// --- Productos ---
var FORM_ENTRY_AGUA_VENTA    = 'entry.413894251';
var FORM_ENTRY_AGUA_REGALO   = 'entry.159366848';
var FORM_ENTRY_CERV_VENTA    = 'entry.718981743';
var FORM_ENTRY_CERV_REGALO   = 'entry.67930836';
var FORM_ENTRY_COCA_VENTA    = 'entry.800397190';
var FORM_ENTRY_COCA_REGALO   = 'entry.1237405773';
 
var PRECIO_AGUA    = 1500;
var PRECIO_CERVEZA = 5000;
var PRECIO_COCA    = 3000;
 
// --- WhatsApp automático (opcional). Lee la guía antes de activarlo ---
var USAR_WHATSAPP   = false;     // ponlo en true solo cuando tengas las API keys
var WA_APIKEY_DUENO = '';        // tu API key, si quieres avisos a tu propio WhatsApp
// =================================================
 
var HOJA_CONFIG   = 'Config';
var HOJA_HORARIOS = 'Horarios';
var HOJA_RESERVAS = 'Reservas';
var HOJA_BLOQUEOS = 'Bloqueos';
var HOJA_BARBEROS = 'Barberos';
var COL_EVENTO    = 15;   // columna donde se guarda el ID del evento de calendario
var COL_AVISO     = 16;   // columna con el botón "Avisar al barbero"
var COL_VALOR_PROD = 28;  // Registro col AB: $ de productos vendidos (va 100% a la barbería)
var DIAS = ['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo'];
 
// ---------- Punto de entrada web ----------
function doGet(e) {
  var p  = e && e.parameter ? e.parameter.p : null;
  var id = e && e.parameter ? e.parameter.c : null;
  var tpl, titulo;

  if (p === 'dashboard') {
    tpl = HtmlService.createTemplateFromFile('Dashboard');
    tpl.reservaId = '';
    titulo = NEGOCIO + ' · Dashboard';
  } else if (id) {
    tpl = HtmlService.createTemplateFromFile('Cancelar');
    tpl.reservaId = id;
    titulo = 'Cancelar cita · ' + NEGOCIO;
  } else {
    tpl = HtmlService.createTemplateFromFile('Reservar');
    tpl.reservaId = '';
    titulo   = 'Reservar en ' + NEGOCIO;
  }

  var out = tpl.evaluate()
    .setTitle(titulo)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  // OJO: addMetaTag() solo acepta viewport y las de "web app" (apple-mobile-
  // web-app-capable, etc.). Etiquetas og:* lanzan "La metaetiqueta que
  // especificaste no se admite" y tumban la página. La tarjeta para compartir
  // (título, descripción, imagen) la da la landing aparte
  // (riifbarberco.github.io/reservas), que redirige a esta página.

  return out;
}

function include(nombre) {
  return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}
 
// ---------- Utilidades ----------
function libro() { return SpreadsheetApp.getActiveSpreadsheet(); }
 
/**
 * Zona horaria de la HOJA. Un valor leído de la hoja hay que formatearlo con
 * la zona de la hoja, no con la del script. Si no coinciden, todas las horas
 * se corren y las citas se solapan.
 */
function tzLibro() {
  try { return libro().getSpreadsheetTimeZone() || TZ; } catch (e) { return TZ; }
}
 
function aHHMM(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, tzLibro(), 'HH:mm');
  }
  var s = String(v).trim();
  var m = s.match(/^(\d{1,2})[:.](\d{2})/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : s;
}
 
function aISO(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, tzLibro(), 'yyyy-MM-dd');
  }
  return String(v).trim().slice(0, 10);
}
 
function aMin(hhmm) {
  var m = String(hhmm).match(/(\d{1,2})[:.](\d{2})/);
  if (!m) return null;
  var min = parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
  return isNaN(min) ? null : min;
}
 
function aTexto(min) {
  var h = Math.floor(min / 60), m = min % 60;
  return ('0' + h).slice(-2) + ':' + ('0' + m).slice(-2);
}
 
function diaSemana(iso) {
  var p = iso.split('-');
  var d = new Date(parseInt(p[0],10), parseInt(p[1],10) - 1, parseInt(p[2],10));
  return DIAS[(d.getDay() + 6) % 7];
}
 
function hoyISO() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function ahoraMin() { return aMin(Utilities.formatDate(new Date(), TZ, 'HH:mm')); }
 
/** Convierte fecha ISO + hora HH:mm en un Date real en la zona horaria correcta */
function aFechaHora(iso, hhmm) {
  return Utilities.parseDate(iso + ' ' + hhmm, TZ, 'yyyy-MM-dd HH:mm');
}
 
function minutosDelDia(fecha) {
  return aMin(Utilities.formatDate(fecha, TZ, 'HH:mm'));
}
 
function duracionTexto(m) {
  var h = Math.floor(m / 60), r = m % 60;
  if (h && r) return h + ' h ' + r + ' min';
  if (h) return h + ' h';
  return r + ' min';
}
 
function pesos(n) {
  return '$' + Number(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}
 
// ---------- Lectura de configuración ----------
function leerConfig() {
  var h = libro().getSheetByName(HOJA_CONFIG);
  if (!h) throw new Error('No existe la hoja Config.');
 
  var barberos = [];
  h.getRange('A5:B10').getValues().forEach(function (f) {
    if (String(f[0]).trim()) {
      barberos.push({ nombre: String(f[0]).trim(), comision: Number(f[1]) || 0 });
    }
  });
 
  var servicios = [];
  h.getRange('D5:F24').getValues().forEach(function (f) {
    if (String(f[0]).trim() && Number(f[2]) > 0) {
      servicios.push({
        nombre: String(f[0]).trim(),
        precio: Number(f[1]) || 0,
        duracion: Number(f[2])
      });
    }
  });
 
  return { barberos: barberos, servicios: servicios };
}

// Cómo se unen los nombres cuando una cita tiene varios servicios.
var SEPARADOR_COMBO = ' + ';

/**
 * Junta uno o varios servicios de Config en un solo "combo": nombre unido,
 * precio y duración sumados. Lo usan crearReserva, crearCitaDashboard y
 * modificarCitaDashboard para que una cita pueda tener más de un servicio.
 *
 * @param {Object} cfg     - resultado de leerConfig()
 * @param {string|string[]} nombres - un nombre o varios
 * @returns {Object} { ok, nombre, precio, duracion } o { ok:false, error }
 */
function _combinarServicios_(cfg, nombres) {
  nombres = (Array.isArray(nombres) ? nombres : [nombres])
    .map(function (n) { return String(n || '').trim(); })
    .filter(Boolean);
  if (!nombres.length) return { ok: false, error: 'Elige al menos un servicio.' };

  var elegidos = [];
  for (var i = 0; i < nombres.length; i++) {
    var s = cfg.servicios.filter(function (x) { return x.nombre === nombres[i]; })[0];
    if (!s) return { ok: false, error: '"' + nombres[i] + '" ya no está disponible. Recarga la página.' };
    elegidos.push(s);
  }
  return {
    ok: true,
    nombre: elegidos.map(function (s) { return s.nombre; }).join(SEPARADOR_COMBO),
    precio: elegidos.reduce(function (a, s) { return a + s.precio; }, 0),
    duracion: elegidos.reduce(function (a, s) { return a + s.duracion; }, 0)
  };
}

/**
 * Escribe en una fila del Registro las fórmulas de Comisión (M), Pago al
 * barbero (N) y Neto barbería (O). SIEMPRE en notación A1 con setFormula():
 * la referencia a una columna entera de OTRA hoja (Config!$A:$A) no se puede
 * expresar de forma fiable con setFormulaR1C1() — Apps Script no traduce bien
 * el atajo "columna sin fila" (Config!C2) al pasarlo a R1C1, y la celda queda
 * con el texto sin interpretar → sale #ERROR!
 *
 * Separador ";" y NO ",": esta hoja tiene configuración regional en español,
 * donde el separador de argumentos de fórmulas es punto y coma. setFormula()
 * no traduce el separador — si se le pasa una fórmula con comas, Sheets no
 * la puede interpretar y también queda en #ERROR! ("Error de análisis de
 * fórmula"), aunque la sintaxis sea válida en inglés.
 */
function _aplicarFormulasComision_(hg, fila) {
  _asegurarColValorProductos_(hg);
  hg.getRange(fila, 13).setFormula(
    '=IF(L' + fila + '="Atendido";IFERROR(I' + fila + '*INDEX(Config!$B:$B;MATCH(C' + fila + ';Config!$A:$A;0));0);0)');
  hg.getRange(fila, 14).setFormula(
    '=IF(L' + fila + '="Atendido";M' + fila + '+IF(J' + fila + '="";0;J' + fila + ');0)');
  // Neto barbería = servicio − comisión + productos vendidos (col AB). Los
  // productos no pasan por la comisión: son 100% de la barbería.
  hg.getRange(fila, 15).setFormula(
    '=IF(L' + fila + '="Atendido";I' + fila + '-M' + fila + '+N(AB' + fila + ');0)');
}

/**
 * Deja lista la col AB (28) "Valor productos" del Registro. La fórmula de Neto
 * la referencia, y si la cuadrícula tiene menos columnas daría #REF!.
 * Se revisa una sola vez por ejecución.
 */
var _colValorProdOk_ = false;
function _asegurarColValorProductos_(hg) {
  if (_colValorProdOk_) return;
  if (hg.getMaxColumns() < COL_VALOR_PROD) {
    hg.insertColumnsAfter(hg.getMaxColumns(), COL_VALOR_PROD - hg.getMaxColumns());
  }
  var enc = hg.getRange(1, COL_VALOR_PROD);
  if (String(enc.getValue()).trim() !== 'Valor productos') {
    enc.setValue('Valor productos')
      .setFontWeight('bold').setBackground('#111111').setFontColor('#FFFFFF');
  }
  _colValorProdOk_ = true;
}

/** { nombreNormalizado: precio } desde la hoja Inventario. */
function _preciosProductos_() {
  var mapa = {};
  _productosCatalogo_().forEach(function (p) { mapa[normalizar(p.nombre)] = p.precio; });
  return mapa;
}

/** $ de los productos VENDIDOS (las cortesías no suman) de un mapa {nombre:{v,r}}. */
function _valorProductos_(prod, precios) {
  precios = precios || _preciosProductos_();
  var total = 0;
  Object.keys(prod || {}).forEach(function (n) {
    total += (Number(prod[n] && prod[n].v) || 0) * (Number(precios[normalizar(n)]) || 0);
  });
  return total;
}

function leerHorarios() {
  var h = libro().getSheetByName(HOJA_HORARIOS);
  if (!h) return {};
  var mapa = {};
  h.getRange(2, 1, Math.max(h.getLastRow() - 1, 0), 4).getValues().forEach(function (f) {
    var barbero = String(f[0]).trim();
    var dia = String(f[1]).trim();
    var abre = aHHMM(f[2]), cierra = aHHMM(f[3]);
    if (!barbero || !dia || !abre || !cierra) return;
    if (aMin(abre) === null || aMin(cierra) === null) return;
    mapa[barbero] = mapa[barbero] || {};
    mapa[barbero][dia] = { abre: abre, cierra: cierra };
  });
  return mapa;
}
 
/** Quita tildes, mayúsculas y espacios para comparar encabezados */
function normalizar(t) {
  return String(t).toLowerCase().trim()
    .replace(/[áàä]/g,'a').replace(/[éèë]/g,'e').replace(/[íìï]/g,'i')
    .replace(/[óòö]/g,'o').replace(/[úùü]/g,'u').replace(/ñ/g,'n')
    .replace(/[^a-z0-9]/g,'');
}
 
/**
 * Contacto y calendario de cada barbero.
 * Busca las columnas POR NOMBRE DE ENCABEZADO, no por posición. Así puedes
 * agregar, mover o quitar columnas en la hoja Barberos sin romper nada.
 * Lo único obligatorio: que exista una columna "Nombre" y que el nombre
 * coincida letra por letra con el de la hoja Config.
 */
function leerBarberos() {
  var h = libro().getSheetByName(HOJA_BARBEROS);
  var mapa = {};
  if (!h || h.getLastRow() < 2) return mapa;
 
  var datos = h.getDataRange().getValues();
  var enc = datos[0].map(normalizar);
 
  function col() {
    for (var i = 0; i < arguments.length; i++) {
      var idx = enc.indexOf(normalizar(arguments[i]));
      if (idx >= 0) return idx;
    }
    return -1;
  }
  function val(fila, idx) { return idx < 0 ? '' : String(fila[idx] || '').trim(); }
 
  var cNombre = col('Nombre', 'Barbero');
  var cCorreo = col('Correo', 'Correo (Gmail)', 'Email', 'Gmail');
  var cWa     = col('WhatsApp', 'Celular', 'Teléfono');
  var cKey    = col('API key WhatsApp', 'API key', 'Apikey');
  var cCal    = col('ID Calendario', 'Calendario', 'IDCalendario');
  var cNotif  = col('Notificar', 'Estado', 'Activo');
  var cPin    = col('PIN', 'Clave', 'Contraseña');
  var cRol    = col('Rol', 'Perfil');
 
  if (cNombre < 0) throw new Error(
    'La hoja Barberos no tiene una columna llamada "Nombre". Revísala.');
 
  for (var f = 1; f < datos.length; f++) {
    var nombre = val(datos[f], cNombre);
    if (!nombre) continue;
    var estado = val(datos[f], cNotif).toLowerCase();
    mapa[nombre] = {
      fila: f + 1,
      colCalendario: cCal + 1,
      nombre: nombre,
      correo: val(datos[f], cCorreo),
      whatsapp: val(datos[f], cWa).replace(/\D/g, ''),
      apikey: val(datos[f], cKey),
      calendario: val(datos[f], cCal),
      notificar: (estado !== 'no' && estado !== 'inactivo' && estado !== 'falso'),
      pin: val(datos[f], cPin).replace(/\D/g, ''),
      colPin: cPin < 0 ? -1 : cPin + 1,
      rol: val(datos[f], cRol)
    };
  }
  return mapa;
}
 
/** Calendario donde se juntan las citas de todos los barberos */
function calendarioGeneral() {
  if (!USAR_CALENDARIO || !CALENDARIO_GENERAL) return null;
  var g = leerBarberos()[CALENDARIO_GENERAL];
  if (!g || !g.calendario) return null;
  try { return CalendarApp.getCalendarById(g.calendario); } catch (e) { return null; }
}
 
// ---------- Ocupación ----------
function ocupacion(barbero, iso, excluirId) {
  var bloques = [];
 
  var hr = libro().getSheetByName(HOJA_RESERVAS);
  if (hr && hr.getLastRow() > 1) {
    hr.getRange(2, 1, hr.getLastRow() - 1, 14).getValues().forEach(function (f) {
      if (aISO(f[2]) !== iso) return;
      if (String(f[5]).trim() !== barbero) return;
      if (String(f[12]).trim() === 'Cancelada') return;
      if (excluirId && String(f[0]).trim() === excluirId) return;   // la cita que se está moviendo
 
      var ini = aMin(aHHMM(f[3]));
      if (ini === null) return;                       // sin hora legible, no se puede bloquear
      var fin = aMin(aHHMM(f[4]));
      if (fin === null || fin <= ini) {               // fin dañado: se reconstruye
        var dur = Number(f[7]) || 0;
        fin = ini + (dur > 0 ? dur : PASO_MIN);
      }
      bloques.push({ ini: ini, fin: fin });
    });
  }
 
  var hb = libro().getSheetByName(HOJA_BLOQUEOS);
  if (hb && hb.getLastRow() > 1) {
    hb.getRange(2, 1, hb.getLastRow() - 1, 5).getValues().forEach(function (f) {
      if (aISO(f[1]) !== iso) return;
      var quien = String(f[0]).trim();
      if (quien && quien !== 'Todos' && quien !== barbero) return;
      var bi = aMin(aHHMM(f[2])), bf = aMin(aHHMM(f[3]));
      if (bi === null || bf === null || bf <= bi) return;
      bloques.push({ ini: bi, fin: bf });
    });
  }
 
  // Eventos puestos a mano en Google Calendar (almuerzo, médico, lo que sea).
  // Los eventos creados por este sistema se ignoran: ya vienen de la hoja Reservas.
  if (USAR_CALENDARIO) {
    var b = leerBarberos()[barbero];
    if (b && b.calendario) {
      try {
        var cal = CalendarApp.getCalendarById(b.calendario);
        if (cal) {
          var desde = aFechaHora(iso, '00:00');
          var hasta = new Date(desde.getTime() + 86400000);
          cal.getEvents(desde, hasta).forEach(function (ev) {
            if (ev.getTag('riif') === '1') return;
            if (ev.isAllDayEvent()) { bloques.push({ ini: 0, fin: 1440 }); return; }
            bloques.push({
              ini: minutosDelDia(ev.getStartTime()),
              fin: minutosDelDia(ev.getEndTime())
            });
          });
        }
      } catch (err) { /* si el calendario falla, la agenda sigue funcionando */ }
    }
  }
 
  return bloques;
}
 
// ---------- API para la página ----------
function getDatosIniciales() {
  var cfg = leerConfig();
  var horarios = leerHorarios();
  var dias = [];
  var base = new Date();
  for (var i = 0; i <= DIAS_ADELANTE; i++) {
    var d = new Date(base.getTime() + i * 86400000);
    dias.push({
      iso: Utilities.formatDate(d, TZ, 'yyyy-MM-dd'),
      dia: DIAS[(d.getDay() + 6) % 7],
      num: Utilities.formatDate(d, TZ, 'd'),
      mes: Utilities.formatDate(d, TZ, 'MMM')
    });
  }
  return {
    negocio: NEGOCIO,
    direccion: DIRECCION,
    servicios: cfg.servicios,
    barberos: cfg.barberos.map(function (b) { return b.nombre; }),
    dias: dias,
    tieneHorarios: Object.keys(horarios).length > 0
  };
}
 
function getCupos(barbero, iso, duracion) {
  duracion = Number(duracion);
  var horarios = leerHorarios();
  var cfg = leerConfig();
  var dia = diaSemana(iso);
  var esHoy = (iso === hoyISO());
  var corte = ahoraMin() + ANTICIPACION_H * 60;
 
  var candidatos = (barbero === '*')
    ? cfg.barberos.map(function (b) { return b.nombre; })
    : [barbero];
 
  var mapa = {}, abreMin = null, cierraMin = null;
 
  candidatos.forEach(function (nom) {
    var h = horarios[nom] && horarios[nom][dia];
    if (!h) return;
    var a = aMin(h.abre), c = aMin(h.cierra);
    abreMin = (abreMin === null) ? a : Math.min(abreMin, a);
    cierraMin = (cierraMin === null) ? c : Math.max(cierraMin, c);
 
    var ocupado = ocupacion(nom, iso);
    for (var t = a; t + duracion <= c; t += PASO_MIN) {
      if (esHoy && t < corte) continue;
      var choca = ocupado.some(function (b) { return t < b.fin && (t + duracion) > b.ini; });
      if (!choca) {
        if (!mapa[t]) mapa[t] = [];
        mapa[t].push(nom);
      }
    }
  });
 
  var cupos = Object.keys(mapa)
    .map(function (t) { return parseInt(t, 10); })
    .sort(function (x, y) { return x - y; })
    .map(function (t) { return { hora: aTexto(t), barberos: mapa[t] }; });
 
  var ocupados = [];
  if (barbero !== '*') {
    ocupados = ocupacion(barbero, iso).map(function (b) {
      return { ini: aTexto(Math.max(b.ini, 0)), fin: aTexto(Math.min(b.fin, 1439)) };
    });
  }
 
  return {
    cupos: cupos,
    abre: abreMin === null ? '09:00' : aTexto(abreMin),
    cierra: cierraMin === null ? '20:00' : aTexto(cierraMin),
    ocupados: ocupados,
    cerrado: (abreMin === null)
  };
}
 
function crearReserva(d) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch (err) { return { ok: false, error: 'El sistema está ocupado. Intenta de nuevo en unos segundos.' }; }
 
  try {
    var cfg = leerConfig();
    var combo = _combinarServicios_(cfg, d.servicios || d.servicio);
    if (!combo.ok) return { ok: false, error: combo.error };

    var nombre = String(d.nombre || '').trim();
    var tel = String(d.telefono || '').replace(/\D/g, '');
    if (nombre.length < 3) return { ok: false, error: 'Escribe tu nombre completo.' };
    if (tel.length < 7) return { ok: false, error: 'Escribe un número de teléfono válido.' };

    var estado = getCupos(d.barbero, d.fecha, combo.duracion);
    var cupo = estado.cupos.filter(function (c) { return c.hora === d.hora; })[0];
    if (!cupo) return { ok: false, error: 'Ese horario se acaba de ocupar. Escoge otro, por favor.' };

    var asignado = cupo.barberos[0];
    if (cupo.barberos.length > 1) {
      var conteo = {};
      cupo.barberos.forEach(function (n) { conteo[n] = ocupacion(n, d.fecha).length; });
      asignado = cupo.barberos.slice().sort(function (a, b) { return conteo[a] - conteo[b]; })[0];
    }

    var finTxt = aTexto(aMin(d.hora) + combo.duracion);
    var id = nuevoIdCita_();

    var r = {
      id: id, fecha: d.fecha, hora: d.hora, fin: finTxt, barbero: asignado,
      servicio: combo.nombre, precio: combo.precio,
      duracion: combo.duracion, nombre: nombre, telefono: tel,
      notas: String(d.notas || '').trim()
    };

    return _guardarCitaNueva_(r, { correoCliente: d.email });

  } catch (err) {
    return { ok: false, error: 'No se pudo guardar la cita: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/** Código único de cita: R + yymmdd + '-' + 5 caracteres. */
function nuevoIdCita_() {
  return 'R' + Utilities.formatDate(new Date(), TZ, 'yyMMdd') + '-' +
         Math.random().toString(36).slice(2, 7).toUpperCase();
}

/**
 * Guarda una cita nueva YA VALIDADA: la escribe en Reservas, crea el evento de
 * calendario, deja el link de aviso, manda los correos, avisa al barbero y la
 * vuelca al Registro. Lo usan crearReserva (web pública) y crearCitaDashboard.
 *
 * @param {Object} r  - { id, fecha, hora, fin, barbero, servicio, precio,
 *                        duracion, nombre, telefono, notas }
 * @param {Object} opciones - { correoCliente }
 */
function _guardarCitaNueva_(r, opciones) {
  opciones = opciones || {};
  var correo = String(opciones.correoCliente || '').trim();

  var h = libro().getSheetByName(HOJA_RESERVAS);
  h.appendRow([
    r.id, new Date(), r.fecha, r.hora, r.fin, r.barbero, r.servicio,
    r.duracion, r.precio, r.nombre, r.telefono,
    correo, 'Confirmada', String(r.notas || '').trim()
  ]);
  var filaNueva = h.getLastRow();

  var idEvento = crearEvento(r);
  if (idEvento) h.getRange(filaNueva, COL_EVENTO).setValue(idEvento);

  var link = linkAvisoBarbero(r);
  if (link) {
    h.getRange(filaNueva, COL_AVISO)
     .setFormula('=HYPERLINK("' + link + '";"Avisar a ' + r.barbero + '")');
  }

  enviarCorreos(r, correo);
  notificarBarbero(r, 'nueva');
  // Ya estamos dentro del candado de crearReserva / crearCitaDashboard:
  // usar la versión sin candado para no bloquearnos a nosotros mismos.
  try { volcarAutomatico_(); } catch (e) {}

  return {
    ok: true,
    reserva: r,
    whatsapp: whatsappDestino(r.barbero),
    destino: WHATSAPP_DESTINO
  };
}

function cancelarReserva(id) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { ok: false, error: 'Intenta de nuevo.' }; }
  try {
    var h = libro().getSheetByName(HOJA_RESERVAS);
    if (!h || h.getLastRow() < 2) return { ok: false, error: 'No encontramos esa cita.' };
    var datos = h.getRange(2, 1, h.getLastRow() - 1, COL_EVENTO).getValues();
 
    for (var i = 0; i < datos.length; i++) {
      if (String(datos[i][0]).trim() !== String(id).trim()) continue;
      if (String(datos[i][12]).trim() === 'Cancelada') {
        return { ok: false, error: 'Esa cita ya estaba cancelada.' };
      }
      h.getRange(i + 2, 13).setValue('Cancelada');
 
      var r = {
        id: datos[i][0], fecha: aISO(datos[i][2]), hora: aHHMM(datos[i][3]),
        fin: aHHMM(datos[i][4]), barbero: String(datos[i][5]).trim(),
        servicio: datos[i][6], duracion: Number(datos[i][7]) || 0, precio: datos[i][8],
        nombre: datos[i][9], telefono: datos[i][10], notas: datos[i][13]
      };
 
      borrarEvento(r.barbero, String(datos[i][14] || '').trim());
      notificarBarbero(r, 'cancelada');
      try { sincronizarCancelaciones(); } catch(e) {}
 
      return { ok: true, reserva: r };
    }
    return { ok: false, error: 'No encontramos esa cita.' };
  } finally {
    lock.releaseLock();
  }
}
 
function verReserva(id) {
  var h = libro().getSheetByName(HOJA_RESERVAS);
  if (!h || h.getLastRow() < 2) return null;
  var datos = h.getRange(2, 1, h.getLastRow() - 1, 14).getValues();
  for (var i = 0; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === String(id).trim()) {
      return {
        id: datos[i][0], fecha: aISO(datos[i][2]), hora: aHHMM(datos[i][3]),
        barbero: datos[i][5], servicio: datos[i][6], nombre: datos[i][9],
        estado: datos[i][12]
      };
    }
  }
  return null;
}
 
// ==================== GOOGLE CALENDAR ====================
 
function calendarioDe(nombreBarbero) {
  if (!USAR_CALENDARIO) return null;
  var b = leerBarberos()[nombreBarbero];
  if (!b || !b.calendario) return null;
  try { return CalendarApp.getCalendarById(b.calendario); } catch (e) { return null; }
}
 
function crearEvento(r) {
  if (!USAR_CALENDARIO) return '';
  try {
    var cal = calendarioDe(r.barbero);
    if (!cal) return '';
 
    var b = leerBarberos()[r.barbero] || {};
    var inicio = aFechaHora(r.fecha, r.hora);
    var fin = new Date(inicio.getTime() + r.duracion * 60000);
 
    var desc =
      'Cliente: ' + r.nombre + '\n' +
      (MOSTRAR_TELEFONO_A_BARBERO && r.telefono
        ? 'Teléfono: ' + r.telefono + '\n' +
          'Escribirle: https://wa.me/57' + String(r.telefono).replace(/^57/, '') + '\n'
        : '') +
      'Servicio: ' + r.servicio + ' (' + duracionTexto(r.duracion) + ')\n' +
      'Valor: ' + pesos(r.precio) + '\n' +
      (r.notas ? 'Notas: ' + r.notas + '\n' : '') +
      'Código: ' + r.id + '\n\n' +
      '— — —\n' +
      '✅ Cerrar este servicio:\n' + linkCierre(r.id);
 
    var opciones = { description: desc, location: DIRECCION };
    if (b.correo && b.notificar) {
      opciones.guests = b.correo;
      opciones.sendInvites = true;
    }
 
    var ev = cal.createEvent(r.servicio + ' — ' + r.nombre, inicio, fin, opciones);
    ev.setTag('riif', '1');
    ev.setTag('reserva', r.id);
    if (RECORDATORIO_MIN > 0) ev.addPopupReminder(RECORDATORIO_MIN);
    var ids = ev.getId();
 
    // Copia en el calendario general, sin invitados para no duplicar correos
    var gen = calendarioGeneral();
    if (gen) {
      var evg = gen.createEvent(r.barbero + ' · ' + r.servicio + ' — ' + r.nombre,
                                inicio, fin, { description: desc, location: DIRECCION });
      evg.setTag('riif', '1');
      evg.setTag('reserva', r.id);
      ids += '|' + evg.getId();
    }
    return ids;
 
  } catch (err) {
    // Si el calendario falla, la cita YA quedó guardada en la hoja. No se pierde nada.
    console.error('Calendario: ' + err.message);
    return '';
  }
}
 
function borrarEvento(nombreBarbero, idEvento) {
  if (!USAR_CALENDARIO || !idEvento) return;
  var partes = String(idEvento).split('|');
  var cals = [calendarioDe(nombreBarbero), calendarioGeneral()];
  for (var i = 0; i < partes.length; i++) {
    if (!partes[i] || !cals[i]) continue;
    try {
      var ev = cals[i].getEventById(partes[i]);
      if (ev) ev.deleteEvent();
    } catch (err) { console.error('Borrar evento: ' + err.message); }
  }
}

/**
 * Marca en Google Calendar que un servicio ya se cerró: cambia el color, le
 * pone un prefijo al título y reemplaza el bloque del link al Form por una
 * línea con el resultado. Recorre los dos eventos (calendario del barbero y
 * calendario general) igual que borrarEvento.
 *
 * Si Calendar falla no pasa nada: el cierre ya quedó guardado en la hoja.
 *
 * @param {string} nombreBarbero
 * @param {string} idEvento - "idBarbero|idGeneral" tal como lo guarda crearEvento
 * @param {string} estado   - Atendido / No asistió / Cancelado
 * @param {string} resumen  - texto corto con servicio, total, método y propina
 */
function marcarEventoCerrado_(nombreBarbero, idEvento, estado, resumen) {
  if (!USAR_CALENDARIO || !idEvento) return;

  var atendido = (String(estado).trim() === 'Atendido');
  var prefijo = atendido ? '✅ ' : '❌ ';
  var color = atendido ? CalendarApp.EventColor.GREEN : CalendarApp.EventColor.GRAY;

  var partes = String(idEvento).split('|');
  var cals = [calendarioDe(nombreBarbero), calendarioGeneral()];

  for (var i = 0; i < partes.length; i++) {
    if (!partes[i] || !cals[i]) continue;
    try {
      var ev = cals[i].getEventById(partes[i]);
      if (!ev) continue;

      // Quitar el prefijo anterior para que volver a cerrar no los apile.
      var titulo = String(ev.getTitle()).replace(/^[✅❌]\s*/, '');
      ev.setTitle(prefijo + titulo);
      try { ev.setColor(color); } catch (e) {}

      // La descripción llevaba el link al Form de cierre; ya no sirve.
      var desc = String(ev.getDescription() || '').split('— — —')[0].replace(/\s+$/, '');
      ev.setDescription(desc + '\n\n— — —\nCerrado: ' + estado +
                        (resumen ? ' · ' + resumen : ''));
    } catch (err) {
      console.error('Marcar evento cerrado: ' + err.message);
    }
  }
}
 
// ==================== NOTIFICACIONES ====================
 
function notificarBarbero(r, tipo) {
  var b = leerBarberos()[r.barbero];
  if (!b || !b.notificar) return;
 
  var esNueva  = (tipo === 'nueva');
  var esModif  = (tipo === 'modificada');
  var cuando = textoFecha(r.fecha) + ' a las ' + r.hora;
  var titulo = esNueva ? 'Nueva cita: ' + cuando
             : esModif ? 'Cita cambiada: ' + cuando
             : 'CANCELADA: ' + cuando;
  var intro  = esNueva ? 'Te agendaron una cita.'
             : esModif ? 'Te cambiaron una cita. Estos son los datos nuevos:'
             : 'Se canceló una cita tuya.';

  var cuerpo =
    intro + '\n\n' +
    'Cuándo: ' + cuando + ' (' + r.hora + ' a ' + r.fin + ')\n' +
    'Servicio: ' + r.servicio + ' · ' + duracionTexto(r.duracion) + '\n' +
    'Cliente: ' + r.nombre + '\n' +
    (MOSTRAR_TELEFONO_A_BARBERO && r.telefono
      ? 'Teléfono: ' + r.telefono + '\n' +
        'Escribirle: https://wa.me/57' + String(r.telefono).replace(/^57/, '') + '\n'
      : '') +
    'Valor: ' + pesos(r.precio) + '\n' +
    (r.notas ? 'Notas: ' + r.notas + '\n' : '') +
    '\nCódigo: ' + r.id;
 
  if (b.correo) {
    try { MailApp.sendEmail(b.correo, titulo, cuerpo); } catch (e) {}
  }
 
  var etiqWa = esNueva ? 'NUEVA CITA' : esModif ? 'CITA CAMBIADA' : 'CITA CANCELADA';
  if (USAR_WHATSAPP && b.whatsapp && b.apikey) {
    enviarWhatsApp(b.whatsapp, b.apikey,
      etiqWa + '\n' +
      cuando + '\n' + r.servicio + '\n' +
      r.nombre + (r.telefono ? ' — ' + r.telefono : '') + '\n' + pesos(r.precio));
  }

  if (USAR_WHATSAPP && WHATSAPP && WA_APIKEY_DUENO) {
    enviarWhatsApp(WHATSAPP, WA_APIKEY_DUENO,
      (esNueva ? 'Nueva cita' : esModif ? 'Cita cambiada' : 'Cita cancelada') + ' · ' + r.barbero + '\n' +
      cuando + ' · ' + r.servicio + '\n' + r.nombre + (r.telefono ? ' — ' + r.telefono : ''));
  }
}
 
/**
 * Arma el mensaje que TÚ le mandas al barbero, y el enlace de WhatsApp
 * que lo abre ya escrito. Un toque tuyo y sale. El barbero se entera por ti,
 * no por el cliente.
 */
function linkAvisoBarbero(r) {
  var b = leerBarberos()[r.barbero];
  if (!b || !b.whatsapp || b.whatsapp.length < 10) return '';
 
  var texto =
    r.barbero + ', te agendé una cita.\n\n' +
    r.servicio + '\n' +
    textoFecha(r.fecha) + '\n' +
    r.hora + ' a ' + r.fin + ' (' + duracionTexto(r.duracion) + ')\n' +
    'Cliente: ' + r.nombre + '\n' +
    (MOSTRAR_TELEFONO_A_BARBERO && r.telefono ? 'Teléfono: ' + r.telefono + '\n' : '') +
    'Valor: ' + pesos(r.precio) + '\n' +
    (r.notas ? 'Nota: ' + r.notas + '\n' : '') +
    'Código: ' + r.id;
 
  return 'https://wa.me/' + b.whatsapp + '?text=' + encodeURIComponent(texto);
}
 
/**
 * Construye el link del form de cierre con el código de la cita prellenado.
 * El barbero lo toca desde el evento de Calendar y solo marca estado y método.
 */
function linkCierre(id) {
  return FORM_CIERRE_URL + '?' +
    FORM_ENTRY_CODIGO + '=' + encodeURIComponent(id) + '&' +
    FORM_ENTRY_ESTADO + '=Atendido&' +
    FORM_ENTRY_METODO + '=Efectivo';
}
 
/**
 * Decide a qué número va el botón de WhatsApp de la confirmación.
 * Si se pidió 'barbero' pero ese barbero no tiene número cargado en la hoja
 * Barberos, cae al número del negocio en vez de dejar el botón roto.
 */
function whatsappDestino(nombreBarbero) {
  if (WHATSAPP_DESTINO === 'barbero') {
    var b = leerBarberos()[nombreBarbero];
    if (b && b.whatsapp && b.whatsapp.length >= 10) return b.whatsapp;
  }
  return WHATSAPP;
}
 
/**
 * Envía un WhatsApp usando CallMeBot.
 * Servicio gratuito de terceros. Cada destinatario lo activa una vez.
 * Si falla, el correo y el calendario siguen llegando igual.
 */
function enviarWhatsApp(numero, apikey, texto) {
  try {
    var url = 'https://api.callmebot.com/whatsapp.php?phone=' + encodeURIComponent(numero) +
              '&text=' + encodeURIComponent(texto) +
              '&apikey=' + encodeURIComponent(apikey);
    UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  } catch (e) { console.error('WhatsApp: ' + e.message); }
}
 
function enviarCorreos(r, emailCliente) {
  var fechaTxt = textoFecha(r.fecha);
 
  if (emailCliente && emailCliente.indexOf('@') > 0) {
    var url = ScriptApp.getService().getUrl() + '?c=' + r.id;
    var cuerpo =
      'Hola ' + r.nombre + ',\n\n' +
      'Tu cita en ' + NEGOCIO + ' quedó confirmada.\n\n' +
      'Servicio: ' + r.servicio + '\n' +
      'Cuándo: ' + fechaTxt + ' a las ' + r.hora + '\n' +
      'Con: ' + r.barbero + '\n' +
      'Valor: ' + pesos(r.precio) + '\n' +
      'Dónde: ' + DIRECCION + '\n\n' +
      'Código de tu cita: ' + r.id + '\n' +
      'Si no puedes venir, cancela aquí y le damos el cupo a alguien más:\n' + url + '\n\n' +
      'Nos vemos.\n' + NEGOCIO;
    try {
      MailApp.sendEmail(emailCliente, 'Cita confirmada en ' + NEGOCIO + ' — ' + fechaTxt, cuerpo);
    } catch (e) {}
  }
 
  if (EMAIL_AVISO && EMAIL_AVISO.indexOf('@') > 0) {
    try {
      var link = linkAvisoBarbero(r);
      MailApp.sendEmail(EMAIL_AVISO, 'Nueva cita: ' + fechaTxt + ' ' + r.hora + ' — ' + r.barbero,
        r.nombre + ' (' + r.telefono + ')\n' + r.servicio + ' · ' + pesos(r.precio) + '\n' +
        fechaTxt + ' ' + r.hora + '–' + r.fin + '\nCon ' + r.barbero + '\nCódigo: ' + r.id +
        (link ? '\n\n— — —\nAvisarle a ' + r.barbero + ' por WhatsApp (abre el chat con el ' +
                'mensaje ya escrito, solo le das enviar):\n' + link : ''));
    } catch (e) {}
  }
}
 
function textoFecha(iso) {
  var meses = ['enero','febrero','marzo','abril','mayo','junio','julio',
               'agosto','septiembre','octubre','noviembre','diciembre'];
  var p = iso.split('-');
  return diaSemana(iso) + ' ' + parseInt(p[2],10) + ' de ' + meses[parseInt(p[1],10) - 1];
}
 
// ==================== INSTALACIÓN Y ACTUALIZACIÓN ====================
 
/**
 * Corre esta función UNA VEZ después de pegar la v2.
 * Es segura: si algo ya existe, no lo toca ni lo duplica.
 */
function actualizarSistema() {
  var lb = libro();
  var mensajes = [];
 
  // 1. Columna del ID de evento, AL FINAL de Reservas (nunca en el medio)
  var hr = lb.getSheetByName(HOJA_RESERVAS);
  if (hr) {
    if (String(hr.getRange(1, COL_EVENTO).getValue()).trim() !== 'ID Evento') {
      hr.getRange(1, COL_EVENTO).setValue('ID Evento')
        .setFontWeight('bold').setBackground('#111111').setFontColor('#FFFFFF');
      mensajes.push('· Se agregó la columna "ID Evento" a la hoja Reservas.');
    }
    if (String(hr.getRange(1, COL_AVISO).getValue()).trim() !== 'Avisar al barbero') {
      hr.getRange(1, COL_AVISO).setValue('Avisar al barbero')
        .setFontWeight('bold').setBackground('#111111').setFontColor('#FFFFFF');
      hr.setColumnWidth(COL_AVISO, 150);
      mensajes.push('· Se agregó la columna "Avisar al barbero" a la hoja Reservas.');
    }
    hr.getRange('D2:E').setNumberFormat('@');
    var hbq = lb.getSheetByName(HOJA_BLOQUEOS);
    if (hbq) hbq.getRange('C2:D').setNumberFormat('@');
    var hhq = lb.getSheetByName(HOJA_HORARIOS);
    if (hhq) hhq.getRange('C2:D').setNumberFormat('@');
 
    var tzL = tzLibro();
    if (tzL !== TZ) {
      mensajes.push('· OJO: la hoja está en ' + tzL + ' y el script en ' + TZ + '. ' +
                    'Ponlas iguales en Archivo → Configuración → Zona horaria.');
    }
  } else {
    mensajes.push('· OJO: no existe la hoja Reservas. Corre primero configurarHojas.');
  }
 
  // 2. Hoja Barberos
  var hb = lb.getSheetByName(HOJA_BARBEROS);
  if (!hb) {
    hb = lb.insertSheet(HOJA_BARBEROS);
    hb.appendRow(['Nombre','Correo (Gmail)','WhatsApp','API key WhatsApp','ID Calendario','Notificar']);
    hb.setFrozenRows(1);
    hb.getRange('A1:F1').setFontWeight('bold').setBackground('#111111').setFontColor('#FFFFFF');
    hb.setColumnWidth(1, 160); hb.setColumnWidth(2, 220); hb.setColumnWidth(3, 130);
    hb.setColumnWidth(4, 150); hb.setColumnWidth(5, 300); hb.setColumnWidth(6, 90);
    hb.getRange('C:D').setNumberFormat('@');
    mensajes.push('· Se creó la hoja Barberos.');
  }

  // 2b. Columnas PIN y Rol para el Dashboard (se agregan al final si faltan)
  var encHb = hb.getRange(1, 1, 1, hb.getLastColumn()).getValues()[0].map(normalizar);
  if (encHb.indexOf(normalizar('PIN')) < 0) {
    var colPin = hb.getLastColumn() + 1;
    hb.getRange(1, colPin).setValue('PIN')
      .setFontWeight('bold').setBackground('#111111').setFontColor('#FFFFFF');
    hb.getRange(2, colPin, Math.max(hb.getMaxRows() - 1, 1), 1).setNumberFormat('@');
    hb.setColumnWidth(colPin, 80);
    mensajes.push('· Se agregó la columna "PIN" a la hoja Barberos ' +
                  '(ponle un PIN de 4 dígitos a cada persona para el Dashboard).');
  }
  encHb = hb.getRange(1, 1, 1, hb.getLastColumn()).getValues()[0].map(normalizar);
  if (encHb.indexOf(normalizar('Rol')) < 0) {
    var colRol = hb.getLastColumn() + 1;
    hb.getRange(1, colRol).setValue('Rol')
      .setFontWeight('bold').setBackground('#111111').setFontColor('#FFFFFF');
    hb.setColumnWidth(colRol, 90);
    mensajes.push('· Se agregó la columna "Rol" a la hoja Barberos ' +
                  '(escribe "Dueño" en tu fila para ver el panel completo).');
  }

  // 3. Sincronizar nombres desde Config, sin borrar lo ya escrito
  var existentes = leerBarberos();
  var nuevos = 0;
  leerConfig().barberos.forEach(function (b) {
    if (!existentes[b.nombre]) {
      hb.appendRow([b.nombre, '', '', '', '', 'Sí']);
      nuevos++;
    }
  });
  if (nuevos) mensajes.push('· Se agregaron ' + nuevos + ' barbero(s) a la hoja Barberos.');

  // 3b. Fila del calendario GENERAL de la barbería (donde caen las citas de
  // todos). Sin esta fila, calendarioGeneral() devuelve null y las citas solo
  // van al calendario del barbero.
  if (USAR_CALENDARIO && CALENDARIO_GENERAL && !existentes[CALENDARIO_GENERAL]) {
    hb.appendRow([CALENDARIO_GENERAL, '', '', '', '', 'No']);
    mensajes.push('· Se agregó la fila del calendario general "' + CALENDARIO_GENERAL +
                  '" a la hoja Barberos.');
  }

  // 4. Calendarios
  if (USAR_CALENDARIO) {
    var creados = configurarCalendarios();
    if (creados.length) mensajes.push('· Calendarios creados: ' + creados.join(', '));
  }
 
  SpreadsheetApp.getUi().alert(
    'Sistema actualizado a v2\n\n' +
    (mensajes.length ? mensajes.join('\n') : 'Ya estaba todo al día.') +
    '\n\nSIGUIENTE PASO:\nAbre la hoja Barberos y escribe el correo Gmail de cada barbero. ' +
    'Sin correo no le llegan ni las invitaciones de calendario ni los avisos.'
  );
}
 
/** Crea un calendario por barbero si aún no existe y guarda su ID */
function configurarCalendarios() {
  var hb = libro().getSheetByName(HOJA_BARBEROS);
  if (!hb) return [];
  var barberos = leerBarberos();
  var creados = [];
 
  var nombresReales = leerConfig().barberos.map(function (x) { return x.nombre; });
 
  Object.keys(barberos).forEach(function (nombre) {
    if (nombresReales.indexOf(nombre) < 0 && nombre !== CALENDARIO_GENERAL) return;
    var b = barberos[nombre];
    var cal = null;
 
    if (b.calendario) {
      try { cal = CalendarApp.getCalendarById(b.calendario); } catch (e) { cal = null; }
    }
    if (!cal) {
      var titulo = PREFIJO_CALENDARIO + nombre;
      var iguales = CalendarApp.getCalendarsByName(titulo);
      cal = iguales.length ? iguales[0]
          : CalendarApp.createCalendar(titulo, {
              summary: 'Citas de ' + nombre + ' en ' + NEGOCIO,
              color: CalendarApp.Color.BROWN,
              timeZone: TZ
            });
      hb.getRange(b.fila, b.colCalendario > 0 ? b.colCalendario : 5).setValue(cal.getId());
      creados.push(titulo);
    }
  });
  return creados;
}
 
function configurarHojas() {
  var lb = libro();
  var cfg = leerConfig();
 
  var hr = lb.getSheetByName(HOJA_RESERVAS);
  if (!hr) {
    hr = lb.insertSheet(HOJA_RESERVAS);
    hr.appendRow(['ID','Creado','Fecha','Hora inicio','Hora fin','Barbero','Servicio',
                  'Duración (min)','Precio','Cliente','Teléfono','Email','Estado','Notas',
                  'ID Evento','Avisar al barbero']);
    hr.getRange('D:E').setNumberFormat('@');
    hr.getRange('C:C').setNumberFormat('yyyy-mm-dd');
    hr.setFrozenRows(1);
    hr.getRange('A1:P1').setFontWeight('bold').setBackground('#111111').setFontColor('#FFFFFF');
  }
 
  var hb = lb.getSheetByName(HOJA_BLOQUEOS);
  if (!hb) {
    hb = lb.insertSheet(HOJA_BLOQUEOS);
    hb.appendRow(['Barbero','Fecha','Desde','Hasta','Motivo']);
    hb.getRange('C:D').setNumberFormat('@');
    hb.getRange('B:B').setNumberFormat('yyyy-mm-dd');
    hb.setFrozenRows(1);
    hb.getRange('A1:E1').setFontWeight('bold').setBackground('#111111').setFontColor('#FFFFFF');
    hb.getRange('A3').setValue('Usa "Todos" en Barbero para cerrar el local completo ese rato.');
  }
 
  var hh = lb.getSheetByName(HOJA_HORARIOS);
  if (!hh) {
    hh = lb.insertSheet(HOJA_HORARIOS);
    hh.appendRow(['Barbero','Día','Abre','Cierra']);
    hh.getRange('C:D').setNumberFormat('@');
    var filas = [];
    cfg.barberos.forEach(function (b) {
      DIAS.forEach(function (d) {
        if (d === 'Domingo') return;
        filas.push([b.nombre, d, '09:00', '20:00']);
      });
    });
    if (filas.length) hh.getRange(2, 1, filas.length, 4).setValues(filas);
    hh.setFrozenRows(1);
    hh.getRange('A1:D1').setFontWeight('bold').setBackground('#111111').setFontColor('#FFFFFF');
  }
 
  actualizarSistema();
}
 
// ==================== UTILIDADES DEL DÍA A DÍA ====================
 
function volcarAlRegistro(iso) {
  iso = iso || hoyISO();
  var hr = libro().getSheetByName(HOJA_RESERVAS);
  var hg = libro().getSheetByName('Registro');
  if (!hr || !hg || hr.getLastRow() < 2) return 0;
 
  var datos = hr.getRange(2, 1, hr.getLastRow() - 1, 14).getValues();
  // Buscar la primera fila vacía en columna A del Registro (no getLastRow que
  // cuenta filas con fórmulas vacías y tira los datos al final)
  var colA = hg.getRange('A2:A3001').getValues();
  var fila = 2;
  for (var fi = 0; fi < colA.length; fi++) {
    if (!colA[fi][0] || String(colA[fi][0]).trim() === '') { fila = fi + 2; break; }
  }
  var n = 0;
  var DIAS = ['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo'];
 
  datos.forEach(function (f) {
    if (aISO(f[2]) !== iso) return;
    if (String(f[12]).trim() === 'Cancelada') return;
 
    // Calcular los valores auxiliares directo, sin depender de fórmulas del Excel
    var fechaDate  = f[2];    // Date de la fecha
    var horaTexto  = aHHMM(f[3]);
    var diaSem     = (fechaDate instanceof Date)
                     ? DIAS[(fechaDate.getDay() + 6) % 7]
                     : diaSemana(aISO(f[2]));
    var horaBloque = horaTexto ? parseInt(horaTexto.split(':')[0], 10) : '';
    var anio       = (fechaDate instanceof Date) ? fechaDate.getFullYear()
                     : parseInt(aISO(f[2]).slice(0,4), 10);
    var mes        = (fechaDate instanceof Date) ? fechaDate.getMonth() + 1
                     : parseInt(aISO(f[2]).slice(5,7), 10);
    var duracion   = Number(f[7]) || 0;   // columna Duración (min) de Reservas
 
    // Columnas exactas del Registro:
    // 1:Fecha 2:Hora 3:Barbero 4:Cliente 5:Teléfono 6:Servicio
    // 7:Valor 8:Desc 9:Total 10:Propina 11:Método 12:Estado
    // 13:Comisión 14:Pago barbero 15:Neto 16:Notas
    // 17:Año 18:Mes 19:Día sem 20:Hora bloque 21:Duración
    // Escribir fecha como objeto Date para que Google Sheets la trate
    // igual que C3 del Cierre diario (ambos Date = SUMIFS funciona)
    var partesFecha = iso.split('-');
    var fechaReal = new Date(
      parseInt(partesFecha[0],10),
      parseInt(partesFecha[1],10) - 1,
      parseInt(partesFecha[2],10)
    );
    hg.getRange(fila, 1).setValue(fechaReal);      // Fecha como Date real
    hg.getRange(fila, 1).setNumberFormat('DD/MM/YYYY');
    hg.getRange(fila, 2).setValue(horaTexto);      // Hora
    hg.getRange(fila, 3).setValue(f[5]);           // Barbero
    hg.getRange(fila, 4).setValue(f[9]);           // Cliente
    hg.getRange(fila, 5).setValue(f[10]);          // Teléfono
    var precio = Number(f[8]) || 0;   // precio de la reserva
 
    hg.getRange(fila, 6).setValue(f[6]);           // Servicio
    hg.getRange(fila, 7).setValue(precio);         // Valor servicio
    // Col 8: Descuento — lo llenas tú si aplica, por defecto 0
    hg.getRange(fila, 8).setValue(0);
    hg.getRange(fila, 9).setValue(precio);         // Total cobrado (precio - descuento, ajusta si hubo descuento)
    // Col 10: Propina — la llenas tú
    hg.getRange(fila, 10).setValue(0);
    // Col 11: Método de pago — lo llenas tú
    // Col 12: Estado — lo llenas tú (Atendido / No asistió / Cancelado)
    hg.getRange(fila, 12).setValue('Pendiente');
    // Cols 13-15: se calculan al momento de cerrar según Estado
    // El código las deja en fórmula para que se actualicen solas cuando
    // cambies el Estado a "Atendido"
    _aplicarFormulasComision_(hg, fila);
    hg.getRange(fila, 16).setValue('Cita ' + f[0]); // Notas
    // Columnas auxiliares
    hg.getRange(fila, 17).setValue(anio);
    hg.getRange(fila, 18).setValue(mes);
    hg.getRange(fila, 19).setValue(diaSem);
    hg.getRange(fila, 20).setValue(horaBloque);
    hg.getRange(fila, 21).setValue(duracion);
 
    fila++; n++;
  });
  return n;
}
 
function volcarHoy() {
  var hoy = hoyISO();
  var ui = SpreadsheetApp.getUi();
 
  // Mostrar qué fecha va a usar para que puedas corregirla si la zona horaria
  // no está configurada aún
  var resp = ui.prompt(
    'Pasar citas al Registro',
    'Fecha a pasar — acepta DD/MM/YYYY o YYYY-MM-DD (ej: 07/08/2026 o 2026-08-07):',
    ui.ButtonSet.OK_CANCEL
  );
  resp.getResponseText; // forzar evaluación
  if (resp.getSelectedButton() !== ui.Button.OK) return;
 
  var rawFecha = resp.getResponseText().trim() || hoy;
  var iso;
  // Acepta DD/MM/YYYY o YYYY-MM-DD
  var mDMY = rawFecha.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  var mYMD = rawFecha.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (mDMY) {
    iso = mDMY[3] + '-' + ('0'+mDMY[2]).slice(-2) + '-' + ('0'+mDMY[1]).slice(-2);
  } else if (mYMD) {
    iso = mYMD[1] + '-' + ('0'+mYMD[2]).slice(-2) + '-' + ('0'+mYMD[3]).slice(-2);
  } else {
    ui.alert('Formato de fecha inválido.\nUsa DD/MM/YYYY (ej: 07/08/2026) o YYYY-MM-DD (ej: 2026-08-07).');
    return;
  }
 
  var n = volcarAlRegistro(iso);
  ui.alert('Se pasaron ' + n + ' citas del ' + iso + ' al Registro.\n\n' +
    (n > 0
      ? 'Falta que completes estado, método de pago y propina en cada fila.'
      : 'No había citas Confirmadas ese día en la hoja Reservas.\n\n' +
        'Verifica que la fecha coincida con la que aparece en la columna Fecha de Reservas.'));
}
 
/** Prueba: manda una notificación de mentira al primer barbero configurado */
function probarNotificaciones() {
  var barberos = leerBarberos();
  var nombre = Object.keys(barberos)[0];
  if (!nombre) { SpreadsheetApp.getUi().alert('No hay barberos en la hoja Barberos.'); return; }
 
  notificarBarbero({
    id: 'PRUEBA', fecha: hoyISO(), hora: '15:00', fin: '15:45',
    barbero: nombre, servicio: 'Corte de Cabello (PRUEBA)', duracion: 45,
    precio: 35000, nombre: 'Cliente de prueba', telefono: '3001234567',
    notas: 'Esto es una prueba. No se agendó ninguna cita real.'
  }, 'nueva');
 
  var b = barberos[nombre];
  SpreadsheetApp.getUi().alert(
    'Notificación de prueba enviada a ' + nombre + '\n\n' +
    'Correo: ' + (b.correo || 'NO CONFIGURADO') + '\n' +
    'WhatsApp: ' + (USAR_WHATSAPP && b.whatsapp && b.apikey ? b.whatsapp : 'apagado o sin API key') + '\n' +
    'Calendario: ' + (b.calendario ? 'listo' : 'NO CONFIGURADO')
  );
}
 
/**
 * Revisa que la hoja Barberos esté bien armada y que los nombres coincidan
 * con Config. Es lo primero que hay que correr cuando "no pasa nada".
 */
function revisarConfiguracion() {
  var problemas = [], ok = [];
  var barberos = leerBarberos();
 
  leerConfig().barberos.forEach(function (b) {
    var d = barberos[b.nombre];
    if (!d) {
      problemas.push('✗ "' + b.nombre + '" está en Config pero NO en la hoja Barberos ' +
                     '(o el nombre no coincide letra por letra).');
      return;
    }
    var faltan = [];
    if (!d.correo) faltan.push('correo');
    if (!d.calendario) faltan.push('calendario');
    if (faltan.length) problemas.push('△ ' + b.nombre + ': falta ' + faltan.join(' y '));
    else ok.push('✓ ' + b.nombre + ' — correo y calendario listos');
  });
 
  if (CALENDARIO_GENERAL) {
    var g = barberos[CALENDARIO_GENERAL];
    if (!g || !g.calendario) {
      problemas.push('△ No encuentro el calendario general "' + CALENDARIO_GENERAL + '".');
    } else {
      ok.push('✓ Calendario general "' + CALENDARIO_GENERAL + '" listo');
    }
  }
 
  SpreadsheetApp.getUi().alert(
    'Revisión de la configuración\n\n' +
    (ok.length ? ok.join('\n') + '\n' : '') +
    (problemas.length ? '\n' + problemas.join('\n') : '\nTodo en orden.')
  );
}
 
/**
 * Muestra qué está leyendo el sistema para un barbero y una fecha.
 * Es la herramienta para entender por qué una hora aparece o no aparece.
 */
function diagnosticoAgenda() {
  var ui = SpreadsheetApp.getUi();
  var cfg = leerConfig();
  if (!cfg.barberos.length) { ui.alert('No hay barberos en Config.'); return; }
 
  var rf = ui.prompt('Diagnóstico de agenda',
    'Fecha a revisar (formato 2026-08-12):', ui.ButtonSet.OK_CANCEL);
  if (rf.getSelectedButton() !== ui.Button.OK) return;
  var iso = rf.getResponseText().trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) { ui.alert('Formato de fecha inválido.'); return; }
 
  var tzL = tzLibro();
  var txt = 'Zona horaria de la hoja: ' + tzL + '\n' +
            'Zona horaria del script: ' + TZ + '\n' +
            (tzL === TZ ? 'Coinciden. Bien.\n' : '¡NO COINCIDEN! Las horas se van a correr.\n') +
            '\nFecha: ' + iso + ' (' + diaSemana(iso) + ')\n';
 
  var horarios = leerHorarios();
  cfg.barberos.forEach(function (b) {
    var h = horarios[b.nombre] && horarios[b.nombre][diaSemana(iso)];
    txt += '\n— ' + b.nombre + ' —\n' +
           'Horario: ' + (h ? h.abre + ' a ' + h.cierra : 'SIN HORARIO ese día') + '\n';
    var oc = ocupacion(b.nombre, iso);
    txt += 'Ocupado: ' + (oc.length
      ? oc.map(function (x) { return aTexto(x.ini) + '-' + aTexto(x.fin); }).join(', ')
      : 'nada') + '\n';
  });
 
  ui.alert('Diagnóstico de agenda', txt, ui.ButtonSet.OK);
}
 
function menuPersonalizado() {
  SpreadsheetApp.getUi()
    .createMenu(NEGOCIO)
    .addItem('Pasar citas de hoy al Registro', 'volcarHoy')
    .addSeparator()
    .addItem('Actualizar sistema / crear calendarios', 'actualizarSistema')
    .addItem('Procesar respuestas del form ahora', 'procesarRespuestasForm')
    .addItem('Correr rutina ahora (prueba)', 'correrRutinaAhora')
    .addSeparator()
    .addItem('Activar automatización', 'instalarDisparador')
    .addItem('Apagar automatización', 'quitarDisparadores')
    .addSeparator()
    .addItem('Configurar inventario de productos', 'configurarInventario')
    .addItem('Reponer stock', 'reponerStock')
    .addSeparator()
    .addItem('Revisar configuración', 'revisarConfiguracion')
    .addItem('Diagnóstico de agenda', 'diagnosticoAgenda')
    .addItem('Probar notificaciones', 'probarNotificaciones')
    .addSeparator()
    .addItem('Reparar fórmulas de comisión y productos', 'repararFormulasComision')
    .addItem('Borrar citas canceladas de la hoja', 'borrarCitasCanceladas')
    .addItem('Sincronizar calendario con el Registro', 'sincronizarCalendarioConRegistro')
    .addToUi();
}

/**
 * Limpieza: borra de Reservas las citas "Cancelada" y del Registro las filas
 * "Cancelado" (con su cita). Los eventos de Calendar ya se borraron al cancelar.
 * Si alguna fila cancelada tenía bebidas anotadas, se devuelven al inventario.
 */
function borrarCitasCanceladas() {
  var ui = SpreadsheetApp.getUi();
  var hr = libro().getSheetByName(HOJA_RESERVAS);
  var hg = libro().getSheetByName('Registro');

  var filasRes = [], ids = {};
  if (hr && hr.getLastRow() >= 2) {
    hr.getRange(2, 1, hr.getLastRow() - 1, 13).getValues().forEach(function (f, i) {
      if (String(f[12]).trim() !== 'Cancelada') return;
      filasRes.push(i + 2);
      ids[String(f[0]).trim()] = 1;
    });
  }
  var filasReg = [];
  if (hg && hg.getLastRow() >= 2) {
    hg.getRange(2, 1, hg.getLastRow() - 1, 16).getValues().forEach(function (f, i) {
      var nota = String(f[15] || '').trim();
      var id = nota.indexOf('Cita ') === 0 ? nota.substring(5).trim() : '';
      if (String(f[11]).trim() === 'Cancelado' || (id && ids[id])) filasReg.push(i + 2);
    });
  }

  if (!filasRes.length && !filasReg.length) { ui.alert('No hay citas canceladas en la hoja.'); return; }
  var r = ui.alert('Borrar citas canceladas',
    'Se van a borrar ' + filasRes.length + ' cita(s) canceladas de Reservas y ' +
    filasReg.length + ' fila(s) del Registro. No se puede deshacer.\n\n¿Continuar?',
    ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var hi = libro().getSheetByName('Inventario');
    // De abajo hacia arriba, para que borrar una fila no corra las demás.
    filasReg.sort(function (a, b) { return b - a; }).forEach(function (fila) {
      if (hi) actualizarInventario_(hi, _deltaProductos_({}, _productosDeFila_(hg, fila)));
      hg.deleteRow(fila);
    });
    filasRes.sort(function (a, b) { return b - a; }).forEach(function (fila) { hr.deleteRow(fila); });
  } finally {
    lock.releaseLock();
  }
  ui.alert('Listo: se borraron ' + filasRes.length + ' cita(s) de Reservas y ' +
           filasReg.length + ' fila(s) del Registro.');
}

/**
 * Vuelve a escribir las fórmulas de Comisión/Pago al barbero/Neto (M, N, O)
 * en TODAS las filas con datos del Registro. Es seguro correrla las veces que
 * quieras: no toca ninguna otra columna, solo deja esas tres fórmulas bien.
 * Se necesitó una vez porque una versión vieja del código las dejaba con
 * #ERROR! en vez de calcular.
 */
function repararFormulasComision() {
  var hg = libro().getSheetByName('Registro');
  var ui = SpreadsheetApp.getUi();
  if (!hg || hg.getLastRow() < 2) { ui.alert('El Registro está vacío, no hay nada que reparar.'); return; }

  var ultima = hg.getLastRow();
  var colA = hg.getRange(2, 1, ultima - 1, 1).getValues();
  _asegurarColValorProductos_(hg);
  // Productos (col 22-27) y su valor (col 28), para completar el $ de las
  // filas cerradas antes de que existiera la columna "Valor productos".
  var prodCols = hg.getRange(2, 22, ultima - 1, 7).getValues();
  var precios = _preciosProductos_();
  var n = 0, nv = 0;
  for (var i = 0; i < colA.length; i++) {
    if (!colA[i][0] || String(colA[i][0]).trim() === '') continue;   // fila vacía, se salta
    _aplicarFormulasComision_(hg, i + 2);
    n++;
    if (prodCols[i][6] === '' || prodCols[i][6] === null) {
      var valor = _valorProductos_(_leerProductosRegistro_(prodCols[i].slice(0, 6)), precios);
      if (valor > 0) { hg.getRange(i + 2, COL_VALOR_PROD).setValue(valor); nv++; }
    }
  }
  // Inventario: Stock actual = inicial + entradas − vendidas − cortesías + ajustes
  var hi = libro().getSheetByName('Inventario');
  if (hi && hi.getLastRow() >= 2) {
    _asegurarColAjustes_(hi);
    var nombresInv = hi.getRange(2, 1, hi.getLastRow() - 1, 1).getValues();
    for (var k = 0; k < nombresInv.length; k++) {
      if (!String(nombresInv[k][0]).trim()) continue;
      var fi = k + 2;
      hi.getRange(fi, 6).setFormula(_formulaStock_(fi));
    }
  }
  ui.alert('Listo: se repararon las columnas M, N y O en ' + n + ' fila(s) del Registro.' +
    (nv ? '\nSe completó el valor de productos en ' + nv + ' fila(s).' : ''));
}
 
function onOpen() { menuPersonalizado(); }
/**
 * RIIF — BLOQUE v5 · Automatización del Registro
 *
 * Se PEGA AL FINAL de Código.gs (después de la última llave). No borra ni
 * modifica nada de lo que ya está: solo agrega funciones nuevas.
 *
 * QUÉ HACE
 *  1. volcarAutomatico()      — pasa las citas al Registro solo, sin preguntar
 *                               fecha y sin duplicar nada.
 *  2. procesarRespuestasForm() — lee el form de cierre de los barberos y
 *                               escribe Estado y Método de pago en el Registro.
 *  3. sincronizarCancelaciones() — si una cita ya volcada se cancela después,
 *                               el Registro pasa a "Cancelado" solo.
 *  4. instalarDisparador()    — instala los dos disparadores de tiempo.
 *
 * Las funciones 1, 2 y 3 corren juntas dentro de rutinaAutomatica().
 * Nota: volcarHoy() (el que pide la fecha) sigue existiendo por si algún día
 * necesitas volcar un día viejo a mano.
 */
 
// Minutos de gracia después de la hora de fin antes de volcar la cita.
// 0 = la vuelca apenas termina. 15 = espera 15 minutos.
var GRACIA_MIN = 10;
 
// ==================== 1. VOLCADO AUTOMÁTICO ====================
 
/**
 * Pasa al Registro todas las citas Confirmadas de hoy cuya hora de fin ya pasó
 * y que todavía no estén en el Registro. Idempotente: se puede correr mil veces
 * al día y nunca duplica una fila.
 */
function volcarAutomatico() {
  // Candado propio: lo llaman a la vez el disparador de 15 min y crearReserva.
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return 0; }
  try {
    return volcarAutomatico_();
  } finally {
    lock.releaseLock();
  }
}

function volcarAutomatico_() {
  var hr = libro().getSheetByName(HOJA_RESERVAS);
  var hg = libro().getSheetByName('Registro');
  if (!hr || !hg || hr.getLastRow() < 2) return 0;

  var datos  = hr.getRange(2, 1, hr.getLastRow() - 1, 14).getValues();
  var yaEsta = idsEnRegistro_();            // { 'RIIF-123': fila }

  // primera fila libre de la columna A (o al final si no hay huecos)
  var ultReg = hg.getLastRow();
  var fila   = ultReg + 1;
  if (ultReg >= 2) {
    var colA = hg.getRange(2, 1, ultReg - 1, 1).getValues();
    for (var fi = 0; fi < colA.length; fi++) {
      if (!colA[fi][0] || String(colA[fi][0]).trim() === '') { fila = fi + 2; break; }
    }
  } else {
    fila = 2;
  }

  var n = 0;
  var DIAS_L = ['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo'];
 
  datos.forEach(function (f) {
    var id = String(f[0]).trim();
    if (!id) return;
    if (yaEsta[id]) return;                    // ya está en el Registro
    var isoCita = aISO(f[2]);
    if (!isoCita) return;
    // Las canceladas SÍ entran: se volcan con estado "Cancelado" para que
    // quede la huella de la cita perdida. Si ya estaba, la arregla
    // sincronizarCancelaciones().
    var cancelada = (String(f[12]).trim() === 'Cancelada');
 
    var fechaDate  = f[2];
    var horaTexto  = aHHMM(f[3]);
    var diaSem     = (fechaDate instanceof Date)
                     ? DIAS_L[(fechaDate.getDay() + 6) % 7]
                     : diaSemana(aISO(f[2]));
    var horaBloque = horaTexto ? parseInt(horaTexto.split(':')[0], 10) : '';
    var p          = isoCita.split('-');
    var fechaReal  = new Date(parseInt(p[0],10), parseInt(p[1],10) - 1, parseInt(p[2],10));
    var precio     = Number(f[8]) || 0;
    var duracion   = Number(f[7]) || 0;
 
    hg.getRange(fila, 1).setValue(fechaReal);
    hg.getRange(fila, 1).setNumberFormat('DD/MM/YYYY');
    hg.getRange(fila, 2).setValue(horaTexto);
    hg.getRange(fila, 3).setValue(f[5]);       // Barbero
    hg.getRange(fila, 4).setValue(f[9]);       // Cliente
    hg.getRange(fila, 5).setValue(f[10]);      // Teléfono
    hg.getRange(fila, 6).setValue(f[6]);       // Servicio
    hg.getRange(fila, 7).setValue(precio);
    hg.getRange(fila, 8).setValue(0);
    hg.getRange(fila, 9).setValue(precio);
    hg.getRange(fila, 10).setValue(0);
    hg.getRange(fila, 12).setValue(cancelada ? 'Cancelado' : 'Pendiente');
 
    _aplicarFormulasComision_(hg, fila);

    hg.getRange(fila, 16).setValue('Cita ' + id);   // ← esta es la llave
    hg.getRange(fila, 17).setValue(fechaReal.getFullYear());
    hg.getRange(fila, 18).setValue(fechaReal.getMonth() + 1);
    hg.getRange(fila, 19).setValue(diaSem);
    hg.getRange(fila, 20).setValue(horaBloque);
    hg.getRange(fila, 21).setValue(duracion);
 
    yaEsta[id] = fila;
    fila++; n++;
  });
 
  return n;
}
 
/** Mapa { idCita: numeroDeFila } leyendo la columna Notas del Registro. */
function idsEnRegistro_() {
  var hg = libro().getSheetByName('Registro');
  var mapa = {};
  if (!hg || hg.getLastRow() < 2) return mapa;
  var notas = hg.getRange(2, 16, hg.getLastRow() - 1, 1).getValues();   // col 16 = Notas
  for (var i = 0; i < notas.length; i++) {
    var t = String(notas[i][0] || '').trim();
    if (t.indexOf('Cita ') === 0) mapa[t.substring(5).trim()] = i + 2;
  }
  return mapa;
}
 
// ==================== 2. FORM DE CIERRE → ESTADO ====================
 
/**
 * Lee las respuestas del formulario de cierre y escribe en el Registro el
 * Estado y el Método de pago de cada cita. Marca cada respuesta ya procesada
 * en una columna al final de la hoja de respuestas, para no repetir trabajo.
 */
function procesarRespuestasForm() {
  if (!FORM_SHEET_ID) return 0;
 
  var hf = SpreadsheetApp.openById(FORM_SHEET_ID).getSheets()[0];
  if (hf.getLastRow() < 2) return 0;
 
  var ancho  = hf.getLastColumn();
  var cabeza = hf.getRange(1, 1, 1, ancho).getValues()[0].map(function (x) {
    return normalizar(String(x));
  });
 
  // Ubicar columnas por el texto del encabezado (aguanta cambios de orden)
  function buscar(claves) {
    for (var i = 0; i < cabeza.length; i++) {
      for (var k = 0; k < claves.length; k++) {
        if (cabeza[i].indexOf(claves[k]) >= 0) return i;
      }
    }
    return -1;
  }
  var cCodigo    = buscar(['codigo', 'cita']);
  var cEstado    = buscar(['estado']);
  var cMetodo    = buscar(['metodo', 'pago']);
  var cPropina   = buscar(['propina']);
  var cAguaV     = buscar(['agua vendida', 'aguavendida']);
  var cAguaR     = buscar(['agua cortesia', 'aguacortesia', 'agua regalo', 'aguaregalo']);
  var cCervV     = buscar(['cerveza vendida', 'cervezavendida']);
  var cCervR     = buscar(['cerveza cortesia', 'cervezacortesia', 'cerveza regalo']);
  var cCocaV     = buscar(['coca cola vendida', 'cocacolavendida', 'coca vendida']);
  var cCocaR     = buscar(['coca cola cortesia', 'cocacolacortesia', 'coca cortesia', 'coca regalo']);
 
  if (cCodigo < 0 || cEstado < 0) {
    throw new Error('No encuentro las columnas "Código de cita" y "Estado" en la hoja de respuestas del form.');
  }
 
  // Columna de marca: la primera libre después de todo
  var colMarca = ancho + 1;
  if (String(hf.getRange(1, ancho).getValue()).trim() === 'Procesado') {
    colMarca = ancho;
  } else {
    hf.getRange(1, colMarca).setValue('Procesado').setFontWeight('bold');
  }
 
  var filas = hf.getRange(2, 1, hf.getLastRow() - 1, colMarca).getValues();
  var hg    = libro().getSheetByName('Registro');
  var hi    = libro().getSheetByName('Inventario');
  var mapa  = idsEnRegistro_();
  var n = 0;
 
  for (var i = 0; i < filas.length; i++) {
    if (String(filas[i][colMarca - 1]).trim() === 'OK') continue;
 
    var id = String(filas[i][cCodigo] || '').trim();
    if (!id) continue;
 
    var filaReg = mapa[id];
    if (!filaReg) continue;
 
    var estado = String(filas[i][cEstado] || '').trim();
    if (estado) hg.getRange(filaReg, 12).setValue(estado);
 
    if (cMetodo >= 0) {
      var metodo = String(filas[i][cMetodo] || '').trim();
      if (metodo) hg.getRange(filaReg, 11).setValue(metodo);
    }
    if (cPropina >= 0) {
      var propina = Number(filas[i][cPropina]);
      if (!isNaN(propina) && propina > 0) hg.getRange(filaReg, 10).setValue(propina);
    }
 
    // --- Productos ---
    var aguaV = Number(filas[i][cAguaV >= 0 ? cAguaV : -1]) || 0;
    var aguaR = Number(filas[i][cAguaR >= 0 ? cAguaR : -1]) || 0;
    var cervV = Number(filas[i][cCervV >= 0 ? cCervV : -1]) || 0;
    var cervR = Number(filas[i][cCervR >= 0 ? cCervR : -1]) || 0;
    var cocaV = Number(filas[i][cCocaV >= 0 ? cCocaV : -1]) || 0;
    var cocaR = Number(filas[i][cCocaR >= 0 ? cCocaR : -1]) || 0;
 
    // Productos → texto en la col 22 (el Form solo trae los 3 fijos; se
    // conservan los productos nuevos que ya hubiera puesto el dashboard).
    // Al inventario se aplica SOLO el delta, para no doble-contar.
    var viejoProd = _productosDeFila_(hg, filaReg);
    var nuevoProd = {};
    Object.keys(viejoProd).forEach(function (n) { nuevoProd[n] = viejoProd[n]; });
    nuevoProd['Agua']      = { v: aguaV, r: aguaR };
    nuevoProd['Cerveza']   = { v: cervV, r: cervR };
    nuevoProd['Coca Cola'] = { v: cocaV, r: cocaR };
    ['Agua', 'Cerveza', 'Coca Cola'].forEach(function (n) {
      if (!nuevoProd[n].v && !nuevoProd[n].r) delete nuevoProd[n];
    });

    hg.getRange(filaReg, 22).setValue(_formatearProductos_(nuevoProd));
    if (hg.getMaxColumns() >= 27) hg.getRange(filaReg, 23, 1, 5).clearContent();
    _asegurarColValorProductos_(hg);
    hg.getRange(filaReg, COL_VALOR_PROD).setValue(_valorProductos_(nuevoProd));
    _aplicarFormulasComision_(hg, filaReg);   // Neto con la fórmula que suma productos

    if (hi) actualizarInventario_(hi, _deltaProductos_(nuevoProd, viejoProd));
 
    hf.getRange(i + 2, colMarca).setValue('OK');
    n++;
  }
  return n;
}
 
/**
 * Aplica un DELTA al inventario. Cada producto se busca por nombre en la col A.
 * delta = { 'Agua': {v:+2, r:0}, 'Gaseosa': {v:-1, r:0} }
 * Hoja Inventario: A Producto | B Stock inicial | C Entradas | D Vendidas |
 *                  E Regaladas | F Stock actual (fórmula) | G Valor vendido | H Precio
 *                  I Ajustes (conteo físico: + si sobró, − si faltó)
 */
function actualizarInventario_(hi, delta) {
  if (!hi || !delta) return;
  var ult = hi.getLastRow();
  if (ult < 2) return;
  var filas = hi.getRange(2, 1, ult - 1, 8).getValues();   // A..H

  Object.keys(delta).forEach(function (nombre) {
    var d = delta[nombre] || {};
    var dv = Number(d.v) || 0, dr = Number(d.r) || 0;
    if (dv === 0 && dr === 0) return;

    var idx = -1;
    for (var i = 0; i < filas.length; i++) {
      if (normalizar(filas[i][0]) === normalizar(nombre)) { idx = i; break; }
    }
    if (idx < 0) return;                       // producto que no está en Inventario

    var fila = idx + 2;
    var precio = Number(filas[idx][7]) || 0;   // col H
    hi.getRange(fila, 4).setValue((Number(hi.getRange(fila, 4).getValue()) || 0) + dv);
    hi.getRange(fila, 5).setValue((Number(hi.getRange(fila, 5).getValue()) || 0) + dr);
    hi.getRange(fila, 7).setValue((Number(hi.getRange(fila, 7).getValue()) || 0) + dv * precio);
  });
}

/**
 * Productos vendidos en una cita → texto para el Registro col 22.
 * "Agua:2/1; Coca Cola:1"  =  nombre:vendidas[/cortesía]
 */
function _formatearProductos_(obj) {
  var partes = [];
  Object.keys(obj || {}).forEach(function (nombre) {
    var p = obj[nombre] || {};
    var v = Number(p.v) || 0, r = Number(p.r) || 0;
    if (v === 0 && r === 0) return;
    partes.push(String(nombre).replace(/[:;]/g, ' ').trim() + ':' + v + (r ? '/' + r : ''));
  });
  return partes.join('; ');
}

function _parsearProductos_(texto) {
  var obj = {};
  String(texto || '').split(';').forEach(function (t) {
    t = t.trim();
    if (!t) return;
    var i = t.lastIndexOf(':');
    if (i < 0) return;
    var nombre = t.slice(0, i).trim();
    if (!nombre) return;
    var cant = t.slice(i + 1).split('/');
    obj[nombre] = { v: Number(cant[0]) || 0, r: Number(cant[1]) || 0 };
  });
  return obj;
}

/**
 * Lee los productos de una fila del Registro. `cols` = valores de las columnas
 * 22..27. Acepta el formato nuevo (texto en col 22) y el viejo (6 números).
 */
function _leerProductosRegistro_(cols) {
  var c0 = cols[0];
  if (typeof c0 === 'string' && c0.indexOf(':') >= 0) return _parsearProductos_(c0);
  var o = {};
  [['Agua', 0, 1], ['Cerveza', 2, 3], ['Coca Cola', 4, 5]].forEach(function (m) {
    var v = Number(cols[m[1]]) || 0, r = Number(cols[m[2]]) || 0;
    if (v || r) o[m[0]] = { v: v, r: r };
  });
  return o;
}

/** Lee los productos de una fila del Registro de forma segura (aunque el grid tenga <27 columnas). */
function _productosDeFila_(hg, fila) {
  var w = Math.min(6, hg.getMaxColumns() - 21);
  var arr = (w > 0) ? hg.getRange(fila, 22, 1, w).getValues()[0] : [];
  while (arr.length < 6) arr.push('');
  return _leerProductosRegistro_(arr);
}

/** Delta entre dos mapas de productos {nombre:{v,r}} → {nombre:{v,r}} con las diferencias. */
function _deltaProductos_(nuevo, viejo) {
  nuevo = nuevo || {}; viejo = viejo || {};
  var d = {};
  var nombres = {};
  Object.keys(nuevo).forEach(function (n) { nombres[n] = 1; });
  Object.keys(viejo).forEach(function (n) { nombres[n] = 1; });
  Object.keys(nombres).forEach(function (n) {
    var a = nuevo[n] || { v: 0, r: 0 }, b = viejo[n] || { v: 0, r: 0 };
    var dv = (Number(a.v) || 0) - (Number(b.v) || 0);
    var dr = (Number(a.r) || 0) - (Number(b.r) || 0);
    if (dv || dr) d[n] = { v: dv, r: dr };
  });
  return d;
}
 
// ==================== 3. CANCELACIONES TARDÍAS ====================
 
/**
 * Deja Reservas (lo que pinta el calendario) igual al Registro (lo que sale
 * en el Resumen). El Registro manda, porque es donde se cierra y se cobra:
 *
 *  1. Fila del Registro sin cita en Reservas (borrada a mano, o escrita a mano
 *     en el Registro sin código) → se crea la cita en Reservas.
 *  2. Registro "Atendido"/"No asistió" pero Reservas "Cancelada" → la cita sí
 *     pasó: Reservas vuelve a "Confirmada".
 *  3. Fecha, hora o barbero distintos → Reservas toma los del Registro.
 *
 * En los casos 1-3, si la cita es de hoy en adelante se (re)crea su evento de
 * Google Calendar. Las filas "Cancelado" del Registro no se tocan. No toma el
 * candado: quien la llama ya debe tenerlo.
 * @returns {number} cuántas citas se arreglaron
 */
function sincronizarReservasConRegistro_() {
  var hg = libro().getSheetByName('Registro');
  var hr = libro().getSheetByName(HOJA_RESERVAS);
  if (!hg || !hr || hg.getLastRow() < 2) return 0;

  var reg = hg.getRange(2, 1, hg.getLastRow() - 1, 21).getValues();
  var res = hr.getLastRow() >= 2 ? hr.getRange(2, 1, hr.getLastRow() - 1, COL_EVENTO).getValues() : [];
  var filaRes = {};
  res.forEach(function (f, i) { var id = String(f[0]).trim(); if (id) filaRes[id] = i; });

  var hoy = hoyISO();
  var n = 0;

  reg.forEach(function (f, i) {
    var estado = String(f[11] || '').trim();
    if (estado === 'Cancelado') return;
    var iso = _isoCeldaRegistro_(f[0]);
    var hora = aHHMM(f[1]);
    var barbero = String(f[2] || '').trim();
    if (!iso || !hora || aMin(hora) === null || !barbero) return;   // fila incompleta: no se adivina

    var nota = String(f[15] || '').trim();
    var id = nota.indexOf('Cita ') === 0 ? nota.substring(5).trim() : '';
    var dur = Number(f[20]) || 0;

    // --- 1. No existe en Reservas: crearla ---
    if (!id || filaRes[id] === undefined) {
      var notaVieja = '';
      if (!id) {
        id = nuevoIdCita_();
        notaVieja = nota;   // lo que hubiera escrito a mano pasa a las notas de la cita
        hg.getRange(i + 2, 16).setValue('Cita ' + id);
      }
      if (!dur) {
        var combo = _combinarServicios_(leerConfig(), String(f[5] || '').trim().split(SEPARADOR_COMBO));
        dur = (combo.ok && combo.duracion) ? combo.duracion : PASO_MIN;
      }
      var r = {
        id: id, fecha: iso, hora: hora, fin: aTexto(aMin(hora) + dur), barbero: barbero,
        servicio: String(f[5] || '').trim(), duracion: dur, precio: Number(f[6]) || 0,
        nombre: String(f[3] || '').trim() || 'Sin nombre',
        telefono: String(f[4] || '').replace(/\D/g, ''), notas: notaVieja
      };
      hr.appendRow([r.id, new Date(), r.fecha, r.hora, r.fin, r.barbero, r.servicio,
                    r.duracion, r.precio, r.nombre, r.telefono, '', 'Confirmada', r.notas]);
      if (iso >= hoy) {
        var idEv = crearEvento(r);
        if (idEv) hr.getRange(hr.getLastRow(), COL_EVENTO).setValue(idEv);
      }
      filaRes[id] = -1;   // ya quedó
      n++;
      return;
    }

    var j = filaRes[id];
    if (j < 0) return;
    var fr = res[j];
    var fila = j + 2;
    var cambio = false;

    // --- 2. Cancelada en Reservas pero atendida en el Registro ---
    if (String(fr[12]).trim() === 'Cancelada' && (estado === 'Atendido' || estado === 'No asistió')) {
      hr.getRange(fila, 13).setValue('Confirmada');
      cambio = true;
    }

    // --- 3. Fecha / hora / barbero distintos: manda el Registro ---
    var durRes = Number(fr[7]) || dur || PASO_MIN;
    if (aISO(fr[2]) !== iso || aHHMM(fr[3]) !== hora || String(fr[5]).trim() !== barbero) {
      hr.getRange(fila, 3).setValue(iso);
      hr.getRange(fila, 4).setValue(hora);
      hr.getRange(fila, 5).setValue(aTexto(aMin(hora) + durRes));
      hr.getRange(fila, 6).setValue(barbero);
      cambio = true;
    }

    if (cambio) {
      try { borrarEvento(String(fr[5]).trim(), String(fr[COL_EVENTO - 1] || '').trim()); } catch (e) {}
      var idEv2 = '';
      if (iso >= hoy) {
        idEv2 = crearEvento({
          id: id, fecha: iso, hora: hora, fin: aTexto(aMin(hora) + durRes), barbero: barbero,
          servicio: String(fr[6]).trim(), duracion: durRes, precio: Number(fr[8]) || 0,
          nombre: String(fr[9]).trim(), telefono: String(fr[10]).trim(), notas: String(fr[13]).trim()
        });
      }
      hr.getRange(fila, COL_EVENTO).setValue(idEv2 || '');
      n++;
    }
  });
  return n;
}

/** Desde el menú: corre la sincronización y cuenta qué hizo. */
function sincronizarCalendarioConRegistro() {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var n;
  try { n = sincronizarReservasConRegistro_(); } finally { lock.releaseLock(); }
  SpreadsheetApp.getUi().alert(n
    ? 'Listo: se arreglaron ' + n + ' cita(s). El calendario ya muestra lo mismo que el Registro.'
    : 'Todo cuadra: el calendario ya muestra lo mismo que el Registro.');
}

/**
 * Si una cita ya estaba en el Registro y después se canceló desde la web,
 * el Registro se pone en "Cancelado" (las fórmulas de comisión se van a cero solas).
 */
function sincronizarCancelaciones() {
  var hr = libro().getSheetByName(HOJA_RESERVAS);
  var hg = libro().getSheetByName('Registro');
  if (!hr || !hg || hr.getLastRow() < 2) return 0;
 
  var datos = hr.getRange(2, 1, hr.getLastRow() - 1, 14).getValues();
  var mapa  = idsEnRegistro_();
  var n = 0;
 
  datos.forEach(function (f) {
    if (String(f[12]).trim() !== 'Cancelada') return;
    var filaReg = mapa[String(f[0]).trim()];
    if (!filaReg) return;
    var actual = String(hg.getRange(filaReg, 12).getValue()).trim();
    if (actual === 'Cancelado' || actual === 'Atendido') return;
    hg.getRange(filaReg, 12).setValue('Cancelado');
    n++;
  });
  return n;
}
 
// ==================== 4. RUTINA Y DISPARADORES ====================
 
/** Lo que corre solo cada 15 minutos. No abre ventanas ni pide nada. */
function rutinaAutomatica() {
  try { volcarAutomatico(); }        catch (e) { console.error('volcar: ' + e); }
  try { procesarRespuestasForm(); }  catch (e) { console.error('form: ' + e); }
  try { sincronizarCancelaciones(); }catch (e) { console.error('cancel: ' + e); }
  try {
    var lock = LockService.getScriptLock();
    if (lock.tryLock(10000)) {
      try { sincronizarReservasConRegistro_(); } finally { lock.releaseLock(); }
    }
  } catch (e) { console.error('sync registro: ' + e); }
}
 
/**
 * Instala el disparador cada 15 minutos. Borra primero los que ya existan
 * para no dejar dos corriendo (eso duplicaría trabajo).
 */
function instalarDisparador() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var f = t.getHandlerFunction();
    if (f === 'rutinaAutomatica' || f === 'volcarAutomatico' || f === 'procesarRespuestasForm') {
      ScriptApp.deleteTrigger(t);
    }
  });
 
  ScriptApp.newTrigger('rutinaAutomatica')
    .timeBased()
    .everyMinutes(15)
    .create();
 
  SpreadsheetApp.getUi().alert(
    'Automatización instalada.\n\n' +
    'Cada 15 minutos, sin que hagas nada:\n' +
    '· Las citas que ya terminaron pasan al Registro\n' +
    '· Las respuestas del form escriben Estado y Método de pago\n' +
    '· Las cancelaciones se reflejan en el Registro\n\n' +
    'Ya no tienes que escribir la fecha nunca más.'
  );
}
 
/** Apaga la automatización (por si necesitas volver a lo manual). */
function quitarDisparadores() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'rutinaAutomatica') { ScriptApp.deleteTrigger(t); n++; }
  });
  SpreadsheetApp.getUi().alert('Se quitaron ' + n + ' disparador(es). Volviste al modo manual.');
}
 
/** Corre la rutina ahora mismo y te dice qué hizo. Úsala para probar. */
function correrRutinaAhora() {
  var a = 0, b = 0, c = 0, err = '';
  try { a = volcarAutomatico(); }         catch (e) { err += '\nVolcado: ' + e; }
  try { b = procesarRespuestasForm(); }   catch (e) { err += '\nForm: ' + e; }
  try { c = sincronizarCancelaciones(); } catch (e) { err += '\nCancelaciones: ' + e; }
 
  SpreadsheetApp.getUi().alert(
    'Rutina ejecutada\n\n' +
    '· Citas pasadas al Registro: ' + a + '\n' +
    '· Respuestas del form aplicadas: ' + b + '\n' +
    '· Cancelaciones sincronizadas: ' + c +
    (err ? '\n\nERRORES:' + err : '')
  );
}
 
// ==================== INVENTARIO DE PRODUCTOS ====================
 
/**
 * Crea la hoja Inventario si no existe y la deja lista con stock inicial.
 * También agrega los encabezados de productos al Registro.
 * Corre esta función UNA VEZ desde el menú.
 */
function configurarInventario() {
  var lb = libro();
  var mensajes = [];
 
  // 1. Hoja Inventario
  var hi = lb.getSheetByName('Inventario');
  if (!hi) {
    hi = lb.insertSheet('Inventario');
    mensajes.push('· Se creó la hoja Inventario.');
  }
 
  // Encabezados (col H "Precio" es la fuente del precio de cada producto)
  hi.getRange(1, 1, 1, 8).setValues([[
    'Producto', 'Stock inicial', 'Entradas', 'Vendidas', 'Regaladas', 'Stock actual', 'Valor vendido', 'Precio'
  ]]).setFontWeight('bold').setBackground('#111111').setFontColor('#FFFFFF');

  // Productos base (solo si la fila está vacía). Precio inicial desde las constantes.
  var productos = [
    ['Agua',      24, 0, 0, 0, '', 0, PRECIO_AGUA],
    ['Cerveza',    0, 0, 0, 0, '', 0, PRECIO_CERVEZA],
    ['Coca Cola',  9, 0, 0, 0, '', 0, PRECIO_COCA]
  ];
  productos.forEach(function(p, i) {
    var fila = i + 2;
    if (!hi.getRange(fila, 1).getValue()) {
      hi.getRange(fila, 1, 1, 8).setValues([p]);
    }
    // asegurar el precio en la col H si la fila ya existía sin él
    if (!hi.getRange(fila, 8).getValue() && hi.getRange(fila, 1).getValue()) {
      hi.getRange(fila, 8).setValue(p[7]);
    }
  });

  _asegurarColAjustes_(hi);

  // Fórmula de stock actual + formatos, para todas las filas con producto
  var ult = Math.max(hi.getLastRow(), 4);
  for (var f = 2; f <= ult; f++) {
    if (!hi.getRange(f, 1).getValue()) continue;
    hi.getRange(f, 6).setFormula(_formulaStock_(f));
    hi.getRange(f, 7).setNumberFormat('$#,##0');
    hi.getRange(f, 8).setNumberFormat('$#,##0');
  }

  hi.setColumnWidth(1, 130);
  hi.setFrozenRows(1);

  // 2. En el Registro, una sola columna de productos (col 22, texto)
  var hr = lb.getSheetByName('Registro');
  if (hr) {
    if (String(hr.getRange(1, 22).getValue()).trim() !== 'Productos') {
      hr.getRange(1, 22).setValue('Productos')
        .setFontWeight('bold').setBackground('#111111').setFontColor('#FFFFFF');
      hr.setColumnWidth(22, 200);
    }
    mensajes.push('· Columna "Productos" lista en el Registro.');
  }

  SpreadsheetApp.getUi().alert(
    'Inventario configurado.\n\n' + mensajes.join('\n') +
    '\n\nEl precio de cada producto se edita en la columna "Precio" de la hoja ' +
    'Inventario (o desde el Dashboard). Para agregar un producto nuevo: escribe ' +
    'una fila nueva con Producto, Stock inicial y Precio, y vuelve a correr esta ' +
    'función — o hazlo desde el Dashboard.'
  );
}
 
/** Stock actual = inicial + entradas − vendidas − cortesías + ajustes. */
function _formulaStock_(f) {
  return '=B' + f + '+C' + f + '-D' + f + '-E' + f + '+N(I' + f + ')';
}

/** Deja lista la col I "Ajustes" del Inventario (la usa la fórmula de stock). */
function _asegurarColAjustes_(hi) {
  if (hi.getMaxColumns() < 9) hi.insertColumnsAfter(hi.getMaxColumns(), 9 - hi.getMaxColumns());
  var enc = hi.getRange(1, 9);
  if (String(enc.getValue()).trim() !== 'Ajustes') {
    enc.setValue('Ajustes').setFontWeight('bold').setBackground('#111111').setFontColor('#FFFFFF');
    enc.setNote('Correcciones por conteo físico: + si había más, − si faltaba. ' +
                'Se llena solo con "Ajustar stock" del Dashboard, o si escribes a mano en "Stock actual".');
  }
}

/**
 * Suma `cantidad` a Entradas del producto (búsqueda por nombre en col A).
 * Núcleo sin UI: lo usan el menú y el Dashboard. Devuelve { ok, error }.
 */
function _reponerStock_(producto, cantidad) {
  var hi = libro().getSheetByName('Inventario');
  if (!hi) return { ok: false, error: 'No existe la hoja Inventario. Corre primero "Configurar inventario".' };
  cantidad = Number(cantidad);
  if (isNaN(cantidad) || cantidad === 0) return { ok: false, error: 'Cantidad inválida.' };

  var ult = hi.getLastRow();
  var filas = ult >= 2 ? hi.getRange(2, 1, ult - 1, 1).getValues() : [];
  for (var i = 0; i < filas.length; i++) {
    if (normalizar(filas[i][0]) === normalizar(producto)) {
      var fila = i + 2;
      var actual = Number(hi.getRange(fila, 3).getValue()) || 0;
      hi.getRange(fila, 3).setValue(actual + cantidad);
      return { ok: true, producto: String(filas[i][0]).trim() };
    }
  }
  return { ok: false, error: 'No encontré el producto "' + producto + '" en Inventario.' };
}

/** Agrega una entrada de inventario (reposición de stock) — desde el menú. */
function agregarStockInventario(producto, cantidad) {
  var r = _reponerStock_(producto, cantidad);
  SpreadsheetApp.getUi().alert(r.ok
    ? 'Stock actualizado: +' + cantidad + ' ' + r.producto + '.'
    : r.error);
}
 
/** Menú para reponer stock manualmente. */
function reponerStock() {
  var ui = SpreadsheetApp.getUi();
  var rProd = ui.prompt('Reponer stock', 'Producto (Agua, Cerveza o Coca Cola):', ui.ButtonSet.OK_CANCEL);
  if (rProd.getSelectedButton() !== ui.Button.OK) return;
  var rCant = ui.prompt('Reponer stock', 'Cantidad a agregar:', ui.ButtonSet.OK_CANCEL);
  if (rCant.getSelectedButton() !== ui.Button.OK) return;
  var cant = parseInt(rCant.getResponseText().trim(), 10);
  if (isNaN(cant) || cant <= 0) { ui.alert('Cantidad inválida.'); return; }
  agregarStockInventario(rProd.getResponseText().trim(), cant);
}
