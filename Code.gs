/**
 * QR-MOY — backend (Google Apps Script, Web App)
 * Implementar como: ejecutar como "Yo" · acceso "Cualquier persona".
 *
 * Propiedades del script (Configuración del proyecto > Propiedades del script). NUNCA van en el index.html:
 *   SITE_URL         la dirección de tu página (con / al final)
 *   TILOPAY_USER, TILOPAY_PASS, TILOPAY_KEY, TILOPAY_MERCHANT (opcional), TILOPAY_TEST ("1" solo en pruebas),
 *                    TILOPAY_CURRENCY ("HNL" por defecto; usa "USD" si tu cuenta aún no tiene lempiras activados)
 *   PAYPAL_CLIENT_ID, PAYPAL_SECRET, PAYPAL_ENV ("sandbox" o "live")
 *   SHEET_ID         (opcional; si el script no está dentro de la hoja)
 *   ADMIN_EMAIL      tu correo, donde te llegan los pagos por transferencia por confirmar
 *   ADMIN_KEY        clave larga y secreta (solo tú); va dentro del enlace de aprobación
 *   MANUAL_INSTRUCTIONS  texto que verá el cliente: banco, número de cuenta, titular y número de Tigo Money
 */

// El servidor es la fuente de verdad de precios y duraciones (la página solo los muestra).
const PLANS = {
  free:    { days: 7, scans: 50 },
  monthly: { days: 30,  tier: "basic", price: { HNL: 199,  USD: 6.99 } },    // Emprendedor mensual
  yearly:  { days: 365, tier: "basic", price: { HNL: 1990, USD: 69.99 } },   // Emprendedor anual
  pro:     { days: 30,  tier: "pro",   price: { HNL: 499,  USD: 17.99 } },   // Negocio mensual
  proyear: { days: 365, tier: "pro",   price: { HNL: 4990, USD: 179.99 } }   // Negocio anual
};
const QUOTA = { free: 1, basic: 10, pro: 30 };   // QR dinámicos vigentes por cuenta
const tier_ = p => p === "free" ? "free" : ((PLANS[p] || {}).tier || "free");
const TL_CUR = () => prop("TILOPAY_CURRENCY") || "HNL"; // moneda del cobro con tarjeta (Tilopay)
const PP_CUR = "USD";                                    // PayPal cobra siempre en dólares
const DAY = 864e5;
const QR = ["id", "url", "owner", "email", "plan", "expires", "created", "updated", "scans"];
const PAY = ["provider", "ref", "qr", "plan", "amount", "currency", "status", "txn", "date"];
const SCR = ["id", "json", "updated"];
const LOG = ["fecha", "qr", "plan", "escaneo_n", "resultado"];
const TL = "https://app.tilopay.com";

const prop = k => PropertiesService.getScriptProperties().getProperty(k) || "";
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const iso = ms => new Date(Number(ms)).toISOString();
const sha = s => Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s).map(b => ("0" + (b & 255).toString(16)).slice(-2)).join("");

// ---------- Hojas ----------
function sheet_(name, cols) {
  const ss = prop("SHEET_ID") ? SpreadsheetApp.openById(prop("SHEET_ID")) : SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || (function () { const s = ss.insertSheet(name); s.appendRow(cols); return s; })();
}
function find_(name, cols, field, val, extra) {
  const v = sheet_(name, cols).getDataRange().getValues(); v.shift();
  for (let i = 0; i < v.length; i++) {
    const o = { _row: i + 2 }; cols.forEach((c, j) => o[c] = v[i][j]);
    if (String(o[field]) === String(val) && (!extra || extra(o))) return o;
  }
  return null;
}
function put_(name, cols, o) {
  const s = sheet_(name, cols), vals = cols.map(c => o[c] === undefined ? (c === "scans" ? 0 : "") : o[c]);
  o._row ? s.getRange(o._row, 1, 1, cols.length).setValues([vals]) : s.appendRow(vals);
  if (name === "QR") formulasQR_(s, o._row || s.getLastRow(), o);
}

// ---------- Entradas ----------
function doGet(e) {
  const p = e.parameter || {}, lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    switch (p.action) {
      case "get": return out(apiGet_(p));
      case "save": return out(apiSave_(p));
      case "load": return out(apiLoad_(p));
      case "status": return out(apiStatus_(p));
      case "page_get": return out(apiPageGet_(p));
      case "form_responses": return out(apiFormResponses_(p));
      case "pay": return out(apiPay_(p));
      case "paypal_capture": return out(apiPaypalCapture_(p));
      case "tilopay_confirm": return out(apiTilopayConfirm_(p));
      case "manual_report": return out(apiManualReport_(p));
      case "approve": return ContentService.createTextOutput(apiApprove_(p));
      default: return out({ status: "error" });
    }
  } catch (err) { return out({ status: "error" }); }
  finally { lock.releaseLock(); }
}

// Webhook de Tilopay (se configura solo vía webhook_url al crear cada enlace de pago).
// No se confía en el contenido: siempre se confirma contra la API de Tilopay.
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const d = JSON.parse(e.postData.contents || "{}");
    if (d.action === "save") return out(apiSave_(d));   // guardar QR (POST: permite enviar varias claves de la cuenta)
    if (d.action === "account") return out(apiAccount_(d));
    if (d.action === "report") return out(apiReport_(d));
    if (d.action === "form_submit") return out(apiFormSubmit_(d));   // respuesta de un formulario (pública)
    if (d.action === "page_save") return out(apiPageSave_(d));   // guardar una página digital (cuerpo largo, por eso va por POST)
    if (String(d.code) === "1" && d.tilopayLinkId && d.orderNumber) apiTilopayConfirm_({ link: d.tilopayLinkId, order: d.orderNumber });
  } catch (err) { /* siempre responder 200 */ }
  finally { lock.releaseLock(); }
  return out({ ok: true });
}

// ---------- QR ----------
const owner_ = (r, key) => !!(r && key && String(key).length >= 20 && sha(key) === r.owner);
const state_ = r => ({ status: "success", id: r.id, plan: r.plan, expires: iso(r.expires), paid: r.plan !== "free" && Date.now() < Number(r.expires),
  scans: Number(r.scans) || 0, limit: r.plan === "free" ? PLANS.free.scans : null });

// Registra CADA escaneo (también los bloqueados) en la hoja ESCANEOS; se crea sola la primera vez
function log_(r, n, res) { try { sheet_("ESCANEOS", LOG).appendRow([new Date(), r.id, r.plan, n, res]); } catch (e) {} }

// Cuando un QR gratis llega al límite te avisa a ti y al cliente (si dejó correo)
function avisoLimite_(r) {
  try {
    const lim = PLANS.free.scans;
    if (prop("ADMIN_EMAIL")) MailApp.sendEmail(prop("ADMIN_EMAIL"), "QR-MOY: límite de escaneos alcanzado (" + r.id + ")",
      "El QR " + r.id + " (plan gratis) llegó a " + lim + " escaneos. Desde el siguiente escaneo ya no redirige.\nCorreo del cliente: " + (r.email || "sin correo"));
    if (r.email) MailApp.sendEmail(r.email, "Tu código QR " + r.id + " llegó al límite de la prueba gratuita",
      "Hola,\n\nTu código QR (" + r.id + ") alcanzó los " + lim + " escaneos de la prueba gratuita y dejará de redirigir. Para mantenerlo activo elige un plan en " + prop("SITE_URL") + " (usa el enlace personal de edición que te enviamos).");
  } catch (e) { /* cuota de correo agotada */ }
}

function apiGet_(p) {
  const r = find_("QR", QR, "id", p.id);
  if (!r) return { status: "notfound" };
  const n = Number(r.scans) || 0;
  if (r.plan === "blocked") { log_(r, n, "suspendido"); return { status: "suspended" }; }   // lo pone el administrador en la hoja
  if (Date.now() > Number(r.expires)) { log_(r, n, "vencido"); return { status: "expired", expires: iso(r.expires) }; }
  if (r.plan === "free" && n >= PLANS.free.scans) { log_(r, n, "bloqueado_limite"); return { status: "limit", limit: PLANS.free.scans }; }
  r.scans = n + 1; // cuenta este escaneo
  put_("QR", QR, r);
  log_(r, r.scans, "ok");
  if (r.plan === "free" && r.scans === PLANS.free.scans) avisoLimite_(r);
  const sc = isPremium_(r) ? find_("PANTALLAS", SCR, "id", r.id) : null;   // pantalla personalizada (solo anual / Negocio)
  return { status: "success", url: r.url, expires: iso(r.expires), plan: r.plan, screen: sc ? sc.json : "" };
}

function apiSave_(p) {
  const id = String(p.id || ""), url = String(p.url || ""), email = String(p.email || "").trim();
  if (!/^[\w-]{3,40}$/.test(id) || !url || url.length > 2000) return { status: "error" };
  const bad = checkUrl_(url); if (bad) return { status: "blocked", reason: bad };
  if (!p.key || String(p.key).length < 20) return { status: "unauthorized" };
  let r = find_("QR", QR, "id", id), isNew = false;
  if (!r) {
    // QR nuevo: hereda el plan vigente de la cuenta (mismo dispositivo/clave) y se respeta el cupo
    const own = sha(p.key); let mine = rowsOf_(own);
    (Array.isArray(p.keys) ? p.keys : []).filter(k => typeof k === "string" && k.length >= 20).slice(0, 10).forEach(k => { const o2 = sha(k); if (o2 !== own) mine = mine.concat(rowsOf_(o2)); });   // otras claves de este dispositivo = misma cuenta
    const b = best_(mine), tier = b ? tier_(b.plan) : "free";
    const used = mine.filter(o => Date.now() < Number(o.expires)).length;
    if (used >= QUOTA[tier]) return { status: "quota", limit: QUOTA[tier] };
    r = { id, owner: own, plan: b ? b.plan : "free", expires: b ? Number(b.expires) : Date.now() + PLANS.free.days * DAY, created: Date.now(), email: "" }; isNew = true;
  } else if (!owner_(r, p.key)) return { status: "unauthorized" };
  const emailChanged = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) && email !== r.email;
  if (emailChanged) r.email = email;
  r.url = url; r.updated = Date.now();
  put_("QR", QR, r);
  if (r.email && (isNew || emailChanged)) {
    try {
      MailApp.sendEmail(r.email, "Tu código QR " + id,
        "Hola,\n\nTu código QR (" + id + ") está guardado. Para editarlo desde cualquier dispositivo abre este enlace personal (no lo compartas):\n" +
        prop("SITE_URL") + "?edit=" + id + "&k=" + p.key + "\n\nVigente hasta: " + iso(r.expires));
    } catch (err) { /* cuota de correo agotada */ }
  }
  saveScreen_(r, p.screen);
  return state_(r);
}

function apiLoad_(p) {
  const r = find_("QR", QR, "id", p.id);
  return owner_(r, p.key) ? Object.assign(state_(r), { url: r.url, email: r.email, screen: (find_("PANTALLAS", SCR, "id", r.id) || {}).json || "" }) : { status: "unauthorized" };
}
function apiStatus_(p) {
  const r = find_("QR", QR, "id", p.id);
  return owner_(r, p.key) ? state_(r) : { status: "unauthorized" };
}

// ---------- Cuentas: cupo y plan compartido ----------
function rowsOf_(owner) {
  const v = sheet_("QR", QR).getDataRange().getValues(); v.shift();
  const out = []; v.forEach((x, i) => { if (String(x[2]) === String(owner)) { const o = { _row: i + 2 }; QR.forEach((c, j) => o[c] = x[j]); out.push(o); } });
  return out;
}
const active_ = o => o.plan !== "free" && o.plan !== "blocked" && Date.now() < Number(o.expires);
// el plan vigente de mayor nivel de la cuenta
function best_(rows) {
  const rank = { pro: 2, basic: 1 };
  return rows.filter(active_).sort((a, b) => (rank[tier_(b.plan)] - rank[tier_(a.plan)]) || (Number(b.expires) - Number(a.expires)))[0];
}

// Los pagos suman tiempo y activan el plan en TODOS los QR de la cuenta.
function markPaid_(pr, txn) {
  pr.status = "paid"; pr.txn = txn; put_("PAGOS", PAY, pr);
  const r = find_("QR", QR, "id", pr.qr);
  const base = r.plan !== "free" && Number(r.expires) > Date.now() ? Number(r.expires) : Date.now();
  r.plan = pr.plan; r.expires = base + PLANS[pr.plan].days * DAY; r.updated = Date.now();
  put_("QR", QR, r);
  rowsOf_(r.owner).forEach(o => { if (o.id !== r.id) { o.plan = r.plan; o.expires = r.expires; o.updated = Date.now(); put_("QR", QR, o); } });
  return state_(r);
}

// ---------- Pantalla personalizada al escanear (anual / Negocio) ----------
const isPremium_ = r => ["yearly", "pro", "proyear"].includes(r.plan) && Date.now() < Number(r.expires);
function cleanScreen_(s) {   // el servidor valida todo: nunca se confía en lo que llega de la página
  try { s = typeof s === "string" ? JSON.parse(s) : s; } catch (e) { return null; }
  if (!s || s.mode !== "screen") return { mode: "direct" };
  const col = c => /^#[0-9a-fA-F]{6}$/.test(c || "") ? c : "#1a66d4";
  const url = u => /^https:\/\/[^\s"'<>]{4,300}$/.test(u || "") ? u : "";
  return { mode: "screen", title: String(s.title || "").slice(0, 60), msg: String(s.msg || "").slice(0, 200), btn: String(s.btn || "Continuar").slice(0, 30),
    delay: [0, 3, 5, 8].includes(Number(s.delay)) ? Number(s.delay) : 3, bg: col(s.bg), accent: col(s.accent), logo: url(s.logo) };
}
function saveScreen_(r, raw) {
  if (raw === undefined || !isPremium_(r)) return;      // si el plan no es premium, se ignora
  const c = cleanScreen_(raw); if (!c) return;
  const e = find_("PANTALLAS", SCR, "id", r.id) || { id: r.id };
  e.json = JSON.stringify(c); e.updated = Date.now(); put_("PANTALLAS", SCR, e);
}

// ---------- Páginas digitales (tarjeta, menú, Link in Bio, producto) ----------
// Cada página usa el mismo ID que su QR dinámico: así heredan plan, cupo, vencimiento y conteo de escaneos.
const PAG = ["id", "type", "owner", "data", "updated"];
const PTYPES = ["card", "links", "menu", "product", "form"];
const RESP = ["fecha", "id", "respuestas"];
function apiPageSave_(p) {
  const id = String(p.id || ""), data = String(p.data || "");
  if (!/^[\w-]{3,40}$/.test(id) || PTYPES.indexOf(p.type) < 0 || data.length > 45000) return { status: "error" };
  const r = find_("QR", QR, "id", id);
  if (!owner_(r, p.key)) return { status: "unauthorized" };
  try { JSON.parse(data); } catch (e) { return { status: "error" }; }
  const e = find_("PAGINAS", PAG, "id", id) || { id };
  e.type = p.type; e.owner = r.owner; e.data = data; e.updated = Date.now(); put_("PAGINAS", PAG, e);
  return { status: "success" };
}
function apiPageGet_(p) {
  const e = find_("PAGINAS", PAG, "id", p.id), r = e && find_("QR", QR, "id", p.id);
  if (!e || !r) return { status: "notfound" };
  if (Date.now() > Number(r.expires)) return { status: "expired" };
  return { status: "success", type: e.type, data: e.data, plan: r.plan };
}

// ---------- Cuenta del dispositivo: plan, cupo y lista de mis QR ----------
const rank_ = o => active_(o) ? ({ pro: 2, basic: 1 }[tier_(o.plan)] || 0) : 0;
function apiAccount_(p) {
  const keys = (Array.isArray(p.keys) ? p.keys : []).filter(k => typeof k === "string" && k.length >= 20).slice(0, 10);
  if (!keys.length) return { status: "error" };
  let rows = []; [...new Set(keys.map(sha))].forEach(o => { rows = rows.concat(rowsOf_(o)); });
  const b = best_(rows);
  if (b) rows.forEach(o => {   // el mejor plan vigente se aplica a todos los QR de la cuenta
    if (rank_(b) > rank_(o) || (rank_(b) === rank_(o) && Number(b.expires) > Number(o.expires))) { o.plan = b.plan; o.expires = b.expires; o.updated = Date.now(); put_("QR", QR, o); }
  });
  const tier = b ? tier_(b.plan) : "free";
  return { status: "success", plan: b ? b.plan : "free", expires: b ? iso(b.expires) : null, quota: QUOTA[tier], used: rows.filter(o => Date.now() < Number(o.expires)).length,
    qrs: rows.sort((x, y) => Number(y.created) - Number(x.created)).slice(0, 40).map(o => ({ id: o.id, url: String(o.url || ""), plan: o.plan, expires: iso(o.expires), scans: Number(o.scans) || 0 })) };
}

// ---------- Protección contra abuso: validar destinos y recibir reportes ----------
const SHORTENERS = ["bit.ly", "tinyurl.com", "t.co", "goo.gl", "is.gd", "v.gd", "cutt.ly", "rb.gy", "ow.ly", "shorturl.at", "t.ly", "buff.ly", "lnkd.in", "tiny.cc", "rebrand.ly", "s.id", "bl.ink"];
function checkUrl_(u) {   // devuelve "" si está bien, o el motivo del bloqueo
  if (/^\s*(javascript|data|vbscript|file|blob|about):/i.test(u)) return "esquema_no_permitido";
  if (!/^https?:/i.test(u)) return "";                                   // tel:, mailto:, texto, vCard, Wi-Fi, etc.
  const m = u.match(/^https?:\/\/([^\/?#:@]+)(?::\d+)?(?:[\/?#]|$)/i);
  if (!m) return "enlace_invalido";                                      // incluye el truco usuario@sitio
  const host = m[1].toLowerCase().replace(/^www\./, "");
  if (SHORTENERS.some(s => host === s || host.endsWith("." + s))) return "acortador";
  if (/^\d{1,3}(\.\d{1,3}){3}$/.test(host) || host.indexOf(".") < 0) return "direccion_no_permitida";   // localhost, IP o nombres internos
  return "";
}
const REP = ["fecha", "qr", "motivo"];
function apiReport_(p) {
  const id = String(p.id || "").slice(0, 40), why = String(p.reason || "").trim().slice(0, 300);
  if (!/^[\w-]{3,40}$/.test(id) || !why) return { status: "error" };
  const r = find_("QR", QR, "id", id); if (!r) return { status: "notfound" };
  const s = sheet_("REPORTES", REP), n = s.getDataRange().getValues().filter(v => v[1] === id).length;
  if (n >= 50) return { status: "success" };                             // tope por QR: evita llenar la hoja
  s.appendRow([new Date(), id, why]);
  if (n < 5) try { if (prop("ADMIN_EMAIL")) MailApp.sendEmail(prop("ADMIN_EMAIL"), "QR-MOY: reporte del QR " + id,
    "Motivo: " + why + "\nDestino: " + r.url + "\nCorreo del dueño: " + (r.email || "sin correo") +
    "\n\nPara suspenderlo, escribe  blocked  en la columna plan (E) de ese QR en la hoja QR. Para reactivarlo, devuelve su plan anterior."); } catch (e) { /* cuota de correo */ }
  return { status: "success" };
}

// ---------- Formularios: recibir y consultar respuestas ----------
function apiFormSubmit_(p) {
  const e = find_("PAGINAS", PAG, "id", p.id), r = e && find_("QR", QR, "id", p.id);
  if (!e || e.type !== "form" || !r || Date.now() > Number(r.expires)) return { status: "notfound" };
  let a = p.answers; if (!Array.isArray(a) || a.length < 1 || a.length > 10) return { status: "error" };
  a = a.map(x => [String((x || [])[0] || "").slice(0, 60), String((x || [])[1] || "").slice(0, 500)]);
  const s = sheet_("RESPUESTAS", RESP);
  if (s.getDataRange().getValues().filter(v => v[1] === p.id).length >= 1000) return { status: "full" };   // tope por formulario
  s.appendRow([new Date(), p.id, JSON.stringify(a)]);
  try { if (r.email) MailApp.sendEmail(r.email, "Nueva respuesta en tu formulario " + p.id, a.map(x => x[0] + ": " + x[1]).join("\n")); } catch (err) { /* cuota de correo agotada */ }
  return { status: "success" };
}
function apiFormResponses_(p) {
  const r = find_("QR", QR, "id", p.id);
  if (!owner_(r, p.key)) return { status: "unauthorized" };
  const v = sheet_("RESPUESTAS", RESP).getDataRange().getValues(); v.shift();
  return { status: "success", rows: v.filter(x => x[1] === p.id).slice(-200).reverse().map(x => [x[0] instanceof Date ? x[0].toISOString() : String(x[0]), String(x[2])]) };
}

// ---------- Columnas de seguimiento en la hoja QR ----------
// Detecta qué separador de argumentos acepta tu hoja según su idioma (coma o punto y coma) y lo recuerda
function sep_(s) {
  const saved = prop("FORMULA_SEP"); if (saved) return saved;
  const t = s.getRange("AZ1"); let sep = ",";
  t.setFormula("=MAX(1,2)"); SpreadsheetApp.flush();
  if (t.getValue() !== 2) { t.setFormula("=MAX(1;2)"); SpreadsheetApp.flush(); if (t.getValue() === 2) sep = ";"; }
  t.clear(); PropertiesService.getScriptProperties().setProperty("FORMULA_SEP", sep);
  return sep;
}
// Escribe en cada fila: vence, dias_activo, dias_restantes, escaneos_restantes, estado
function formulasQR_(s, row, o) {
  const tz = s.getParent().getSpreadsheetTimeZone(), r = row, lim = PLANS.free.scans, lim80 = Math.floor(lim * 0.8);
  const sp = sep_(s), f = x => x.replace(/,/g, sp);
  const offH = Math.round(Number(Utilities.formatDate(new Date(), tz, "Z")) / 100);   // desfase de tu zona horaria en horas (Honduras: -6)
  const nowUtc = "(NOW()-(" + offH + "/24))";
  s.getRange(r, 10).setValue(o.expires ? Utilities.formatDate(new Date(Number(o.expires)), tz, "yyyy-MM-dd HH:mm") : "");
  s.getRange(r, 11, 1, 4).setFormulas([[
    f('=IF(G' + r + '="","",ROUNDDOWN(' + nowUtc + '-(G' + r + '/86400000+25569),0))'),
    f('=IF(F' + r + '="","",MAX(0,ROUNDUP((F' + r + '/86400000+25569)-' + nowUtc + ',0)))'),
    f('=IF(E' + r + '="free",MAX(0,' + lim + '-I' + r + '),"ilimitado")'),
    f('=IF(F' + r + '="","",IF(E' + r + '="blocked","SUSPENDIDO",IF(L' + r + '=0,"VENCIDO",IF(E' + r + '="free",IF(I' + r + '>=' + lim + ',"LÍMITE ALCANZADO",IF(I' + r + '>=' + lim80 + ',"CASI EN EL LÍMITE","Activo")),"Activo"))))')
  ]]);
}
// Ejecútala UNA vez a mano: crea los encabezados, los colores del estado y rellena los QR que ya existen
function agregarColumnasQR() {
  const s = sheet_("QR", QR);
  s.getRange(1, 9, 1, 6).setValues([["scans", "vence", "dias_activo", "dias_restantes", "escaneos_restantes", "estado"]]);
  const rg = s.getRange("N2:N"), rule = (t, c) => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t).setBackground(c).setRanges([rg]).build();
  s.setConditionalFormatRules(s.getConditionalFormatRules().concat([rule("LÍMITE ALCANZADO", "#f8b4b4"), rule("CASI EN EL LÍMITE", "#fde68a"), rule("VENCIDO", "#e5e7eb"), rule("SUSPENDIDO", "#fca5a5")]));
  const n = s.getLastRow() - 1;
  if (n < 1) return;
  s.getRange(2, 1, n, QR.length).getValues().forEach((x, i) => {
    const r = {}; QR.forEach((c, j) => r[c] = x[j]);
    if (r.scans === "") s.getRange(i + 2, 9).setValue(0);
    formulasQR_(s, i + 2, r);
  });
}

// ---------- Iniciar pago ----------
function apiPay_(p) {
  const r = find_("QR", QR, "id", p.id), plan = PLANS[p.plan];
  if (!owner_(r, p.key)) return { status: "unauthorized" };
  if (!plan || !plan.price) return { status: "error" };
  const site = prop("SITE_URL"), desc = "QR-MOY " + p.plan + " (" + r.id + ")";

  if (p.provider === "manual") { // transferencia bancaria / Tigo Money, se confirma a mano
    const amount = (plan.price.HNL || 0).toFixed(2), ref = "MOY-" + Utilities.getUuid().replace(/-/g, "").slice(0, 6).toUpperCase();
    put_("PAGOS", PAY, { provider: "manual", ref, qr: r.id, plan: p.plan, amount, currency: "HNL", status: "pending", date: Date.now() });
    return { status: "success", ref, amount, currency: "HNL", instructions: prop("MANUAL_INSTRUCTIONS") };
  }

  if (p.provider === "tilopay") {
    const cur = TL_CUR(), amount = (plan.price[cur] || 0).toFixed(2);
    if (Number(amount) <= 0) return { status: "error" };
    const tok = http_(TL + "/api/v1/login", { apiuser: prop("TILOPAY_USER"), password: prop("TILOPAY_PASS") }).access_token;
    const res = http_(TL + "/api/v1/createLinkPayment", {
      key: prop("TILOPAY_KEY"), amount, currency: cur, reference: r.id + "|" + p.plan, type: 1,
      description: desc, callback_url: site, webhook_url: ScriptApp.getService().getUrl()
    }, tok);
    if (!res.url || !res.id) return { status: "error" };
    put_("PAGOS", PAY, { provider: "tilopay", ref: String(res.id), qr: r.id, plan: p.plan, amount, currency: cur, status: "pending", date: Date.now() });
    return { status: "success", url: res.url };
  }

  if (p.provider === "paypal") {
    const amount = plan.price.USD.toFixed(2);
    const o = pp_("post", "/v2/checkout/orders", {
      intent: "CAPTURE",
      purchase_units: [{ custom_id: r.id + "|" + p.plan, description: desc, amount: { currency_code: PP_CUR, value: amount } }],
      payment_source: { paypal: { experience_context: { return_url: site + "?pp=1", cancel_url: site, user_action: "PAY_NOW", shipping_preference: "NO_SHIPPING" } } }
    }).json;
    const link = (o.links || []).filter(l => l.rel === "payer-action" || l.rel === "approve")[0];
    if (!o.id || !link) return { status: "error" };
    put_("PAGOS", PAY, { provider: "paypal", ref: o.id, qr: r.id, plan: p.plan, amount, currency: PP_CUR, status: "pending", date: Date.now() });
    return { status: "success", url: link.href };
  }
  return { status: "error" };
}

// ---------- Confirmar pago: PayPal ----------
function apiPaypalCapture_(p) {
  const pr = find_("PAGOS", PAY, "ref", p.order, o => o.provider === "paypal");
  if (!pr) return { status: "notfound" };
  if (pr.status === "paid") return state_(find_("QR", QR, "id", pr.qr));
  let res = pp_("post", "/v2/checkout/orders/" + encodeURIComponent(p.order) + "/capture", {});
  if (res.json.status !== "COMPLETED") res = pp_("get", "/v2/checkout/orders/" + encodeURIComponent(p.order));
  const pu = (res.json.purchase_units || [])[0] || {}, cap = ((pu.payments || {}).captures || [])[0] || {};
  const ok = res.json.status === "COMPLETED" && cap.status === "COMPLETED" && cap.amount &&
    cap.amount.currency_code === pr.currency && Number(cap.amount.value) >= Number(pr.amount);
  return ok ? markPaid_(pr, cap.id || p.order) : { status: "pending" };
}

// ---------- Confirmar pago: Tilopay ----------
function apiTilopayConfirm_(p) {
  const order = String(p.order || ""), pr = find_("PAGOS", PAY, "ref", p.link, o => o.provider === "tilopay");
  if (!pr || !order) return { status: "notfound" };
  if (pr.status === "paid") return state_(find_("QR", QR, "id", pr.qr));
  if (find_("PAGOS", PAY, "txn", order, o => o.status === "paid")) return { status: "pending" }; // orderNumber ya usado
  const tok = http_(TL + "/api/v1/login", { apiuser: prop("TILOPAY_USER"), password: prop("TILOPAY_PASS") }).access_token;
  const res = http_(TL + "/api/v1/consult", { key: prop("TILOPAY_KEY"), orderNumber: order, merchantId: prop("TILOPAY_MERCHANT") }, tok);
  const ok = (res.response || []).filter(x =>
    String(x.code) === "1" && String(x.orderNumber) === order && x.currency === pr.currency &&
    Number(x.amount) >= Number(pr.amount) && (prop("TILOPAY_TEST") === "1" || x.environment === "Production")).length > 0;
  return ok ? markPaid_(pr, order) : { status: "pending" };
}

// ---------- Pago manual (transferencia / Tigo Money) ----------
// El cliente avisa que pagó; tú verificas el dinero en tu banco o Tigo Money y apruebas con el enlace del correo.
function apiManualReport_(p) {
  const pr = find_("PAGOS", PAY, "ref", p.ref, o => o.provider === "manual"), r = pr && find_("QR", QR, "id", pr.qr);
  if (!pr || !owner_(r, p.key)) return { status: "unauthorized" };
  if (pr.status === "paid") return state_(r);
  if (pr.status === "pending") {
    pr.status = "reported"; pr.txn = String(p.note || "").slice(0, 120); put_("PAGOS", PAY, pr);
    try {
      MailApp.sendEmail(prop("ADMIN_EMAIL"), "QR-MOY: pago por confirmar " + pr.ref,
        "Referencia: " + pr.ref + "\nQR: " + pr.qr + "\nPlan: " + pr.plan + "\nMonto: L " + pr.amount +
        "\nCorreo del cliente: " + r.email + "\nNota del cliente: " + pr.txn +
        "\n\nPrimero verifica que el dinero ya entró a tu banco o Tigo Money con esa referencia. Si es así, aprueba aquí:\n" +
        ScriptApp.getService().getUrl() + "?action=approve&ref=" + pr.ref + "&admin=" + prop("ADMIN_KEY"));
    } catch (e) { /* cuota de correo agotada */ }
  }
  return { status: "success", reported: true };
}

function apiApprove_(p) {
  if (!prop("ADMIN_KEY") || p.admin !== prop("ADMIN_KEY")) return "No autorizado";
  const pr = find_("PAGOS", PAY, "ref", p.ref, o => o.provider === "manual");
  if (!pr) return "Referencia no encontrada";
  if (pr.status === "paid") return "Este pago ya estaba aprobado";
  const s = markPaid_(pr, "manual:" + pr.ref), r = find_("QR", QR, "id", pr.qr);
  try { if (r.email) MailApp.sendEmail(r.email, "QR-MOY: pago confirmado", "Tu pago fue confirmado. Tu QR " + r.id + " está activo hasta " + s.expires + "."); } catch (e) {}
  return "Plan " + pr.plan + " activado para " + pr.qr + " hasta " + s.expires;
}

// ---------- HTTP ----------
function http_(url, body, token) {
  const r = UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(body),
    headers: token ? { Authorization: "Bearer " + token } : {}, muteHttpExceptions: true });
  try { return JSON.parse(r.getContentText() || "{}"); } catch (e) { return {}; }
}
function pp_(method, path, body) {
  const base = prop("PAYPAL_ENV") === "live" ? "https://api-m.paypal.com" : "https://api-m.sandbox.paypal.com";
  const t = UrlFetchApp.fetch(base + "/v1/oauth2/token", { method: "post", payload: "grant_type=client_credentials",
    headers: { Authorization: "Basic " + Utilities.base64Encode(prop("PAYPAL_CLIENT_ID") + ":" + prop("PAYPAL_SECRET")) }, muteHttpExceptions: true });
  const tok = JSON.parse(t.getContentText() || "{}").access_token;
  const opt = { method, headers: { Authorization: "Bearer " + tok }, muteHttpExceptions: true };
  if (method === "post") { opt.contentType = "application/json"; opt.payload = JSON.stringify(body || {}); }
  const r = UrlFetchApp.fetch(base + path, opt);
  let json = {}; try { json = JSON.parse(r.getContentText() || "{}"); } catch (e) {}
  return { code: r.getResponseCode(), json };
}
