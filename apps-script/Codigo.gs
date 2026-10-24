/**
 * finanzas pe — lector automático de correos bancarios.
 *
 * Corre en tu cuenta de Google cada 10 minutos: lee las alertas de BCP, Scotiabank y
 * Bank of America, las convierte en movimientos y los guarda en esta hoja de cálculo.
 * La app los descarga al abrirse (y guarda aquí su copia de seguridad).
 *
 * Instalación: ver LEEME.md. Resumen: pegar este archivo, ejecutar `instalar`,
 * implementar como aplicación web y ejecutar `verEnlace`.
 */

const CONFIG = {
  ZONA: 'America/Lima',
  REMITENTES: [
    'notificaciones@notificacionesbcp.com.pe',
    'alertas@scotiabank.com.pe',
    'bancadigital@scotiabank.com.pe',
    'bankofamerica.com',
  ],
  DIAS_INICIALES: 180,     // en la primera lectura, cuántos días hacia atrás
  HOJA_MOV: 'Movimientos',
  HOJA_REV: 'Revisar',
  HOJA_RESP: 'Respaldo',
};

const COLS_MOV = ['ID correo', 'Fecha', 'Hora', 'Tipo', 'Descripción', 'Monto', 'Moneda',
  'Origen banco', 'Origen cuenta', 'Origen nombre', 'Destino banco', 'Destino cuenta', 'Destino nombre',
  'Monto destino', 'Moneda destino', 'Confianza', 'Asunto', 'Leído el'];
const COLS_REV = ['ID correo', 'Fecha', 'Remitente', 'Asunto', 'Extracto', 'Enlace', 'Leído el'];

// ─────────────────────────── INSTALACIÓN ───────────────────────────

function instalar() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(CONFIG.ZONA);
  hoja_(CONFIG.HOJA_MOV, COLS_MOV);
  hoja_(CONFIG.HOJA_REV, COLS_REV);
  hoja_(CONFIG.HOJA_RESP, ['Fecha', 'Parte']);
  const hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hoja1 && ss.getSheets().length > 1 && hoja1.getLastRow() === 0) ss.deleteSheet(hoja1);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TOKEN')) props.setProperty('TOKEN', Utilities.getUuid().replace(/-/g, ''));

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'procesarCorreos')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('procesarCorreos').timeBased().everyMinutes(10).create();

  const n = procesarCorreos();
  Logger.log('✓ Instalado. Leí ' + n + ' movimientos nuevos.');
  Logger.log('Siguiente paso: Implementar → Nueva implementación → Aplicación web, y luego ejecuta verEnlace.');
}

/**
 * Muestra el enlace que se pega en la app (Config → Sincronización).
 * Desde el editor, getUrl() a veces devuelve el enlace de prueba (/dev), que solo funciona con tu sesión
 * de Google; la app necesita el de la implementación (/exec).
 */
function verEnlace() {
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  const url = ScriptApp.getService().getUrl() || '';
  if (/\/exec$/.test(url)) {
    Logger.log('Pega esto en la app (Config → Sincronización):\n\n' + url + '?t=' + token);
    return;
  }
  Logger.log('1. Ve a Implementar → Gestionar implementaciones y copia la "URL" de la aplicación web (termina en /exec).\n' +
    '2. Pega en la app esa URL seguida de esto (sin espacios):\n\n?t=' + token +
    '\n\nQueda así: https://script.google.com/macros/s/…/exec?t=' + token);
}

/** Vuelve a leer todo desde cero (por ejemplo, tras mejorar una regla). Borra lo leído, no el respaldo. */
function releerTodo() {
  [CONFIG.HOJA_MOV, CONFIG.HOJA_REV].forEach(n => {
    const h = SpreadsheetApp.getActive().getSheetByName(n);
    // Deja una fila vacía: Google no permite borrar todas las filas bajo el encabezado fijo
    if (h && h.getLastRow() > 2) h.deleteRows(3, h.getLastRow() - 2);
    if (h && h.getLastRow() > 1) h.getRange(2, 1, 1, h.getLastColumn()).clearContent();
  });
  PropertiesService.getScriptProperties().deleteProperty('ULTIMA_LECTURA');
  Logger.log('Leí ' + procesarCorreos() + ' movimientos.');
}

function hoja_(nombre, columnas) {
  const ss = SpreadsheetApp.getActive();
  let h = ss.getSheetByName(nombre);
  if (!h) h = ss.insertSheet(nombre);
  if (h.getLastRow() === 0) {
    h.appendRow(columnas);
    h.getRange(1, 1, 1, columnas.length).setFontWeight('bold');
    h.setFrozenRows(1);
  }
  return h;
}

// ─────────────────────────── LECTURA DE CORREOS ───────────────────────────

function procesarCorreos() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return 0;
  try {
    const inicio = Date.now();
    const props = PropertiesService.getScriptProperties();
    const hMov = hoja_(CONFIG.HOJA_MOV, COLS_MOV);
    const hRev = hoja_(CONFIG.HOJA_REV, COLS_REV);
    const vistos = new Set([...idsDe_(hMov), ...idsDe_(hRev)]);

    // Desde la última lectura (con 3 días de margen) o, la primera vez, DIAS_INICIALES atrás
    const ultima = Number(props.getProperty('ULTIMA_LECTURA')) || 0;
    const desde = ultima ? new Date(ultima - 3 * 864e5) : new Date(Date.now() - CONFIG.DIAS_INICIALES * 864e5);
    const query = 'from:(' + CONFIG.REMITENTES.join(' OR ') + ') after:' + Utilities.formatDate(desde, CONFIG.ZONA, 'yyyy/MM/dd');

    const filasMov = [], filasRev = [];
    const ahora = Utilities.formatDate(new Date(), CONFIG.ZONA, 'yyyy-MM-dd HH:mm');
    for (let start = 0; Date.now() - inicio < 4.5 * 60e3; start += 100) {
      const hilos = GmailApp.search(query, start, 100);
      if (!hilos.length) break;
      hilos.forEach(hilo => hilo.getMessages().forEach(msg => {
        const id = msg.getId();
        if (vistos.has(id)) return;
        vistos.add(id);
        const remitente = msg.getFrom();
        if (!CONFIG.REMITENTES.some(r => remitente.toLowerCase().includes(r))) return;
        const r = parsearCorreo(remitente, msg.getSubject(), msg.getPlainBody(), msg.getDate());
        if (r && r.ignorar) return;
        if (r) {
          filasMov.push([id, r.fecha, r.hora, r.tipo, r.desc, r.monto, r.moneda,
            r.origen.banco, r.origen.cuenta, r.origen.nombre,
            r.destino.banco, r.destino.cuenta, r.destino.nombre,
            r.montoDestino || '', r.monedaDestino || '', r.confianza, msg.getSubject(), ahora]);
        } else {
          filasRev.push([id, Utilities.formatDate(msg.getDate(), CONFIG.ZONA, 'yyyy-MM-dd'), remitente,
            msg.getSubject(), normalizar(msg.getPlainBody()).slice(0, 300),
            'https://mail.google.com/mail/u/0/#all/' + id, ahora]);
        }
      }));
      if (hilos.length < 100) break;
    }
    if (filasMov.length) {
      // Cuentas y montos como texto para que Sheets no les quite ceros ni formato
      hMov.getRange(hMov.getLastRow() + 1, 1, filasMov.length, COLS_MOV.length).setNumberFormat('@').setValues(filasMov);
    }
    if (filasRev.length) hRev.getRange(hRev.getLastRow() + 1, 1, filasRev.length, COLS_REV.length).setValues(filasRev);
    props.setProperty('ULTIMA_LECTURA', String(inicio));
    return filasMov.length;
  } finally {
    lock.releaseLock();
  }
}

function idsDe_(hoja) {
  const n = hoja.getLastRow() - 1;
  return n > 0 ? hoja.getRange(2, 1, n, 1).getValues().map(r => String(r[0])) : [];
}

// ─────────────────────────── INTERPRETACIÓN ───────────────────────────
// Funciones puras: no usan servicios de Google (salvo la fecha), así se pueden probar fuera.

const MONTO = '(S\\/\\.?|US\\$|\\$|USD)\\s?([\\d,]+(?:\\.\\d{1,2})?)';

function normalizar(s) {
  return String(s || '')
    .replace(/ /g, ' ')
    .replace(/\[\]\(#\)/g, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/[|*_]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}
function moneda_(simbolo) { return /S/.test(simbolo) ? 'PEN' : 'USD'; }
function numero_(s) { return Math.round(parseFloat(String(s).replace(/,/g, '')) * 100) / 100; }
function buscar_(t, re) { const m = t.match(re); return m || null; }
function montoDe_(t, etiqueta) {
  const m = t.match(new RegExp(etiqueta + '\\s*:?\\s*' + MONTO, 'i'));
  return m ? { monto: numero_(m[2]), moneda: moneda_(m[1]) } : null;
}
/** Últimos dígitos de la cuenta o tarjeta que sigue a `etiqueta` ("Desde Clasica **** 4076" → "4076", "XXX-XXXXXXXX-0-76" → "076"). */
function sufijo_(t, etiqueta) {
  const m = t.match(new RegExp(etiqueta + '\\s*:?[^\\d]{0,45}?(\\d(?:-?\\d){1,3})(?!\\d)', 'i'));
  return m ? m[1].replace(/-/g, '') : '';
}
/** Nombre de la cuenta antes de los dígitos ("Desde Clasica 4076" → "Clasica"). */
function nombre_(t, etiqueta) {
  const m = t.match(new RegExp(etiqueta + '\\s*:?\\s*([^\\d]{0,40}?)\\s*(?:X{2,}[X-]*)?\\d', 'i'));
  return m ? m[1].replace(/[X-]+$/, '').trim() : '';
}
function comercio_(s) {
  let c = String(s || '').trim().replace(/^(DLC|PYU|MP|IZI|PAYU|SQ|PP|DL|NIUBIZ|CULQI)\s+/i, '');
  const plin = c.match(/^PLIN-(.+)/i);
  if (plin) return 'Plin a ' + titulo_(plin[1]);
  return titulo_(c);
}
function titulo_(s) {
  return String(s).toLowerCase().replace(/(^|[\s\/-])([a-záéíóúñ])/g, (m, a, b) => a + b.toUpperCase()).trim();
}
function cuenta_(banco, cuenta, nombre) { return { banco: banco || '', cuenta: cuenta || '', nombre: nombre || '' }; }
const NADA = cuenta_();
const EFECTIVO = cuenta_('Efectivo', 'efectivo', 'Efectivo');

const IGNORAR = /rechaz|visualizaci|configuraci|estado de cuenta|c[oó]digo de verificaci|dispositivo|actualizaste|viabcp|dep[oó]sito a plazo|cuenta power|bienvenid|encuesta|token|clave|activaci[oó]n|bloqueo|desbloqueo|afiliaci/i;

/**
 * Convierte un correo en movimiento.
 * Devuelve { tipo: 'gasto'|'ingreso'|'transferencia'|'pago', desc, monto, moneda, origen, destino, ... },
 * { ignorar: true } si es un aviso sin dinero, o null si no lo entiende (va a la hoja "Revisar").
 */
function parsearCorreo(remitente, asunto, cuerpo, fecha) {
  const t = normalizar(cuerpo);
  const a = String(asunto || '');
  const de = String(remitente || '').toLowerCase();
  const base = {
    fecha: formatoFecha_(fecha, 'yyyy-MM-dd'), hora: formatoFecha_(fecha, 'HH:mm'),
    origen: NADA, destino: NADA, confianza: 'alta',
  };
  const r = de.includes('bcp') ? bcp_(t, a)
    : de.includes('scotiabank') ? scotiabank_(t, a)
    : de.includes('bankofamerica') ? bofa_(t, a)
    : null;
  if (r) return Object.assign(base, r);
  return IGNORAR.test(a) ? { ignorar: true } : null;
}

function bcp_(t, a) {
  let m;
  // Consumo con tarjeta de crédito o débito
  if ((m = buscar_(t, new RegExp('Realizaste un consumo de ' + MONTO + ' con tu Tarjeta de (Cr[eé]dito|D[eé]bito) BCP en (.+?)\\. Por tu seguridad', 'i')))) {
    const credito = /^cr/i.test(m[3]);
    const suf = sufijo_(t, 'N[uú]mero de Tarjeta de (?:Cr[eé]dito|D[eé]bito)');
    return {
      tipo: 'gasto', desc: comercio_(m[4]), monto: numero_(m[2]), moneda: moneda_(m[1]),
      origen: cuenta_('BCP', suf, (credito ? 'Tarjeta de Crédito' : 'Tarjeta de Débito') + ' BCP'),
    };
  }
  // Retiro en cajero
  if ((m = buscar_(t, new RegExp('Realizaste un retiro de ' + MONTO + ' con tu Tarjeta de D[eé]bito BCP', 'i')))) {
    const cuenta = sufijo_(t, 'Cuenta de cargo') || sufijo_(t, 'N[uú]mero de Tarjeta de D[eé]bito');
    return {
      tipo: 'transferencia', desc: 'Retiro en cajero', monto: numero_(m[2]), moneda: moneda_(m[1]),
      origen: cuenta_('BCP', cuenta, 'Cuenta BCP'), destino: EFECTIVO,
    };
  }
  // Retiro por ventanilla
  if ((m = buscar_(t, new RegExp('retiro de tu cuenta por ventanilla de ' + MONTO, 'i')))) {
    return {
      tipo: 'transferencia', desc: 'Retiro por ventanilla', monto: numero_(m[2]), moneda: moneda_(m[1]),
      origen: cuenta_('BCP', sufijo_(t, 'N[uú]mero de cuenta'), 'Cuenta BCP'), destino: EFECTIVO,
    };
  }
  // Pago de tarjeta propia
  if ((m = buscar_(t, new RegExp('Realizaste un pago a tu tarjeta de ' + MONTO, 'i')))) {
    return {
      tipo: 'pago', desc: 'Pago de tarjeta', monto: numero_(m[2]), moneda: moneda_(m[1]),
      origen: cuenta_('BCP', sufijo_(t, 'Desde'), nombre_(t, 'Desde')),
      destino: cuenta_('BCP', sufijo_(t, 'Pagado a'), nombre_(t, 'Pagado a') || 'Tarjeta de Crédito BCP'),
    };
  }
  // Transferencias enviadas (entre mis cuentas u a otros bancos/terceros)
  if ((m = buscar_(t, new RegExp('Realizaste una transferencia de ' + MONTO + ' desde tu', 'i')))) {
    const recibido = { monto: numero_(m[2]), moneda: moneda_(m[1]) };
    // Si hubo tipo de cambio o comisión, lo cobrado al origen difiere de lo que llega al destino
    const cobrado = montoDe_(t, 'Total cobrado al tipo de cambio') || montoDe_(t, 'Total cobrado') || recibido;
    const enviado = montoDe_(t, 'Monto enviado') || montoDe_(t, 'Monto transferido') || recibido;
    const propias = /entre mis cuentas/i.test(t) || /entre mis cuentas/i.test(a);
    const bancoDestino = titulo_(((buscar_(t, /Banco destino\s*:?\s*(.+?)\s+Moneda/i) || [])[1] || '').replace(/\s+S\.?a\.?a\.?$/i, ''));
    return {
      tipo: 'transferencia',
      desc: propias ? 'Transferencia entre mis cuentas' : 'Transferencia' + (bancoDestino ? ' a ' + bancoDestino : ''),
      monto: cobrado.monto, moneda: cobrado.moneda,
      montoDestino: enviado.monto, monedaDestino: enviado.moneda,
      origen: cuenta_('BCP', sufijo_(t, 'Desde'), nombre_(t, 'Desde')),
      destino: cuenta_(propias ? 'BCP' : (bancoDestino || 'Otro banco'), sufijo_(t, 'Enviado a'), propias ? nombre_(t, 'Enviado a') : 'Otra cuenta'),
    };
  }
  // Yape recibido
  if ((m = buscar_(t, new RegExp('Recibiste un yapeo de ' + MONTO + ' de (.+?)\\. Por tu seguridad', 'i')))) {
    return {
      tipo: 'ingreso', desc: 'Yape de ' + titulo_(m[3]), monto: numero_(m[2]), moneda: moneda_(m[1]),
      destino: cuenta_('BCP', 'yape', 'Yape'),
    };
  }
  // Yape enviado (formato supuesto, aún sin correo de ejemplo)
  if ((m = buscar_(t, new RegExp('(?:Yapeaste|Realizaste un yapeo de) ' + MONTO + ' a (.+?)\\.', 'i')))) {
    return {
      tipo: 'gasto', desc: 'Yape a ' + titulo_(m[3]), monto: numero_(m[2]), moneda: moneda_(m[1]),
      origen: cuenta_('BCP', 'yape', 'Yape'), confianza: 'baja',
    };
  }
  // Depósito recibido
  if ((m = buscar_(t, new RegExp('Te realizaron un dep[oó]sito de ' + MONTO, 'i')))) {
    return {
      tipo: 'ingreso', desc: 'Depósito en cuenta', monto: numero_(m[2]), moneda: moneda_(m[1]),
      destino: cuenta_('BCP', sufijo_(t, 'Cuenta de abono'), 'Cuenta BCP'),
    };
  }
  // Transferencia recibida (formato supuesto, aún sin correo de ejemplo)
  if ((m = buscar_(t, new RegExp('Recibiste una transferencia de ' + MONTO + '(?: de (.+?)\\.)?', 'i')))) {
    return {
      tipo: 'ingreso', desc: m[3] ? 'Transferencia de ' + titulo_(m[3]) : 'Transferencia recibida',
      monto: numero_(m[2]), moneda: moneda_(m[1]),
      destino: cuenta_('BCP', sufijo_(t, '(?:Cuenta de abono|Enviado a|Cuenta destino)'), 'Cuenta BCP'), confianza: 'baja',
    };
  }
  // Pago de servicios
  if (/Pago de servicios/i.test(t) && (m = buscar_(t, new RegExp('Monto total\\s*:?\\s*' + MONTO, 'i')))) {
    const empresa = (buscar_(t, /Empresa\s*:?\s*(.+?)\s+Servicio/i) || [])[1];
    return {
      tipo: 'gasto', desc: empresa ? 'Pago ' + titulo_(empresa) : 'Pago de servicio', monto: numero_(m[2]), moneda: moneda_(m[1]),
      origen: cuenta_('BCP', sufijo_(t, 'Cuenta de origen'), nombre_(t, 'Cuenta de origen')),
    };
  }
  return null;
}

function scotiabank_(t, a) {
  let m;
  // Compra (por internet o presencial) con tarjeta
  if ((m = buscar_(t, new RegExp('Tarjeta de (Cr[eé]dito|D[eé]bito) de Scotiabank que termina en (\\d{4}) por ' + MONTO + ' en (.+?) el \\d{1,2}\\/\\d{1,2}', 'i')))) {
    const credito = /^cr/i.test(m[1]);
    return {
      tipo: 'gasto', desc: comercio_(m[5]), monto: numero_(m[4]), moneda: moneda_(m[3]),
      origen: cuenta_('Scotiabank', m[2], (credito ? 'Tarjeta de Crédito' : 'Tarjeta de Débito') + ' Scotiabank'),
    };
  }
  // Pago de tarjeta de crédito
  if (/constancia de tu pago de Tarjeta de Cr[eé]dito/i.test(t) && (m = buscar_(t, new RegExp('Monto pagado\\s*:?\\s*' + MONTO, 'i')))) {
    return {
      tipo: 'pago', desc: 'Pago de tarjeta', monto: numero_(m[2]), moneda: moneda_(m[1]),
      origen: cuenta_('Scotiabank', sufijo_(t, 'Pagado con'), nombre_(t, 'Pagado con')),
      destino: cuenta_('Scotiabank', sufijo_(t, 'Tarjeta pagada'), nombre_(t, 'Tarjeta pagada') || 'Tarjeta Scotiabank'),
    };
  }
  return null;
}

/** Bank of America: aún no hay correos de ejemplo; reglas generales para sus alertas en inglés. */
function bofa_(t, a) {
  const m = t.match(/\$\s?([\d,]+\.\d{2})/);
  if (!m) return null;
  const suf = (t.match(/ending in (\d{4})/i) || [])[1] || '';
  const deposito = /deposit|received|credited/i.test(a + ' ' + t);
  const lugar = (t.match(/(?:\bat|Where:|Merchant:)\s+(.+?)(?:\s+on\s|\s+Date|\s+When|\.\s|$)/i) || [])[1];
  if (!deposito && !/transaction|purchase|charge|debit|payment|withdrawal/i.test(a + ' ' + t)) return null;
  const cta = cuenta_('Bank of America', suf, 'Bank of America');
  return deposito
    ? { tipo: 'ingreso', desc: 'Depósito Bank of America', monto: numero_(m[1]), moneda: 'USD', destino: cta, confianza: 'baja' }
    : { tipo: 'gasto', desc: lugar ? comercio_(lugar) : 'Compra Bank of America', monto: numero_(m[1]), moneda: 'USD', origen: cta, confianza: 'baja' };
}

function formatoFecha_(d, patron) {
  return Utilities.formatDate(d instanceof Date ? d : new Date(d), CONFIG.ZONA, patron);
}

// ─────────────────────────── API PARA LA APP ───────────────────────────

function doGet(e) {
  if (!autorizado_(e.parameter.t)) return json_({ ok: false, error: 'token' });
  if (e.parameter.accion === 'historico') {
    return json_({ ok: true, cierres: historicos_(String(e.parameter.pares || '').split(',')) });
  }
  if (e.parameter.accion === 'precios') {
    return json_({ ok: true, precios: precios_(String(e.parameter.tickers || '').split(',')) });
  }
  if (e.parameter.accion === 'respaldo') {
    const r = leerRespaldo_();
    return json_({ ok: true, respaldo: r.datos, fecha: r.fecha });
  }
  const hMov = hoja_(CONFIG.HOJA_MOV, COLS_MOV);
  const n = hMov.getLastRow() - 1;
  const filas = n > 0 ? hMov.getRange(2, 1, n, COLS_MOV.length).getDisplayValues() : [];
  const movimientos = filas.map(f => ({
    id: f[0], fecha: f[1], hora: f[2], tipo: f[3], desc: f[4], monto: Number(f[5]), moneda: f[6],
    origen: { banco: f[7], cuenta: f[8], nombre: f[9] },
    destino: { banco: f[10], cuenta: f[11], nombre: f[12] },
    montoDestino: f[13] ? Number(f[13]) : null, monedaDestino: f[14] || null, confianza: f[15],
  }));
  const hRev = hoja_(CONFIG.HOJA_REV, COLS_REV);
  // No se cuentan los avisos sin dinero que quedaron ahí antes de que existiera su regla (ej: activación de tarjeta)
  const nRev = hRev.getLastRow() - 1;
  const sinLeer = nRev > 0 ? hRev.getRange(2, 4, nRev, 1).getValues().filter(r => r[0] && !IGNORAR.test(String(r[0]))).length : 0;
  return json_({ ok: true, movimientos, sinLeer, hoja: SpreadsheetApp.getActive().getUrl(), actualizado: Utilities.formatDate(new Date(), CONFIG.ZONA, 'yyyy-MM-dd HH:mm') });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'json' }); }
  if (!autorizado_(body.t)) return json_({ ok: false, error: 'token' });
  if (body.accion === 'respaldo' && typeof body.datos === 'string') {
    guardarRespaldo_(body.datos);
    return json_({ ok: true });
  }
  return json_({ ok: false, error: 'accion' });
}

function autorizado_(t) {
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  return !!token && t === token;
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Precios de ETFs y acciones con GOOGLEFINANCE (gratis, ~20 min de retraso).
// Cada ticker queda en la pestaña "Precios" con su fórmula; Google la recalcula solo.
// Un ticker nuevo puede devolver "cargando" la primera vez: aparece en la siguiente sincronización.
function precios_(tickers) {
  const h = hoja_('Precios', ['Ticker', 'Precio', 'Moneda', 'Nombre']);
  const pedidos = [...new Set(tickers.map(t => String(t).trim().toUpperCase()).filter(t => /^[A-Z0-9.:\-]{1,20}$/.test(t)))];
  const n = h.getLastRow() - 1;
  const existentes = n > 0 ? h.getRange(2, 1, n, 1).getValues().map(r => String(r[0]).toUpperCase()) : [];
  const nuevos = pedidos.filter(t => !existentes.includes(t));
  if (nuevos.length) {
    const fila = h.getLastRow() + 1;
    h.getRange(fila, 1, nuevos.length, 4).setFormulas(nuevos.map((t, i) => {
      const r = fila + i;
      return ['="' + t + '"', formulaPrecio_(t, r), '=IFERROR(GOOGLEFINANCE(A' + r + ',"currency"),"")', '=IFERROR(GOOGLEFINANCE(A' + r + ',"name"),"")'];
    }));
    SpreadsheetApp.flush();
  }
  // Filas de monedas creadas con la fórmula de acciones (no funciona para CURRENCY:): se corrigen
  const nn = h.getLastRow() - 1;
  if (nn > 0) {
    const forms = h.getRange(2, 1, nn, 2).getFormulas(), vals = h.getRange(2, 1, nn, 1).getValues();
    let arreglo = false;
    vals.forEach((v, i) => {
      const t = String(v[0]).toUpperCase(), r = i + 2;
      if (/^CURRENCY:/.test(t) && forms[i][1] !== formulaPrecio_(t, r)) { h.getRange(r, 2).setFormula(formulaPrecio_(t, r)); arreglo = true; }
    });
    if (arreglo) SpreadsheetApp.flush();
  }
  const m = h.getLastRow() - 1;
  const filas = m > 0 ? h.getRange(2, 1, m, 4).getValues() : [];
  const out = {};
  filas.forEach(f => {
    const t = String(f[0]).toUpperCase();
    if (pedidos.includes(t)) out[t] = typeof f[1] === 'number' && f[1] > 0 ? { precio: f[1], moneda: f[2], nombre: f[3] } : null;
  });
  return out;
}

// Precio de cierre de un ticker en una fecha pasada (el último cierre hasta ese día), con GOOGLEFINANCE.
// Se usa para estimar fondos que siguen a un índice (ej: un portafolio que replica el S&P 500 → VOO).
function historicos_(pares) {
  const h = hoja_('Historico', ['Clave', 'Cierre']);
  const pedidos = [...new Set(pares.map(x => String(x).trim().toUpperCase()).filter(x => /^[A-Z0-9.:\-]{1,20}@\d{4}-\d{2}-\d{2}$/.test(x)))];
  const n = h.getLastRow() - 1;
  const existentes = n > 0 ? h.getRange(2, 1, n, 1).getValues().map(r => String(r[0]).toUpperCase()) : [];
  const nuevos = pedidos.filter(k => !existentes.includes(k));
  if (nuevos.length) {
    const fila = h.getLastRow() + 1;
    h.getRange(fila, 1, nuevos.length, 2).setFormulas(nuevos.map(k => {
      const [t, f] = k.split('@'), [y, m, d] = f.split('-').map(Number);
      const rango = 'GOOGLEFINANCE("' + t + '","close",DATE(' + y + ',' + m + ',' + d + ')-7,DATE(' + y + ',' + m + ',' + d + ')+1)';
      return ['="' + k + '"', '=IFERROR(INDEX(' + rango + ',ROWS(' + rango + '),2),"")'];
    }));
    SpreadsheetApp.flush();
  }
  const m = h.getLastRow() - 1;
  const out = {};
  (m > 0 ? h.getRange(2, 1, m, 2).getValues() : []).forEach(f => {
    const k = String(f[0]).toUpperCase();
    if (pedidos.includes(k)) out[k] = typeof f[1] === 'number' && f[1] > 0 ? f[1] : null;
  });
  return out;
}

// Monedas (CURRENCY:USDPEN) se piden sin atributo; acciones y ETFs con "price"
function formulaPrecio_(t, r) {
  return /^CURRENCY:/.test(t) ? '=IFERROR(GOOGLEFINANCE(A' + r + '),"")' : '=IFERROR(GOOGLEFINANCE(A' + r + ',"price"),"")';
}

// El respaldo se parte en trozos porque una celda admite ~50.000 caracteres
function guardarRespaldo_(datos) {
  const h = hoja_(CONFIG.HOJA_RESP, ['Fecha', 'Parte']);
  // Se vacían las celdas en vez de borrar filas: Google no deja borrar todas las filas bajo el encabezado fijo
  if (h.getLastRow() > 1) h.getRange(2, 1, h.getLastRow() - 1, 2).clearContent();
  const fecha = Utilities.formatDate(new Date(), CONFIG.ZONA, 'yyyy-MM-dd HH:mm');
  const partes = [];
  for (let i = 0; i < datos.length; i += 45000) partes.push([fecha, datos.slice(i, i + 45000)]);
  h.getRange(2, 1, partes.length, 2).setNumberFormat('@').setValues(partes);
}
function leerRespaldo_() {
  const h = hoja_(CONFIG.HOJA_RESP, ['Fecha', 'Parte']);
  const n = h.getLastRow() - 1;
  if (n < 1) return { datos: null, fecha: null };
  const filas = h.getRange(2, 1, n, 2).getDisplayValues().filter(f => f[1]);
  if (!filas.length) return { datos: null, fecha: null };
  return { datos: filas.map(f => f[1]).join(''), fecha: filas[0][0] };
}
