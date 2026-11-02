/**
 * Evaluación de Simulacros VIME · Servidor (Google Apps Script)
 * Guarda establecimientos y simulacros en una hoja de cálculo de su Google Drive,
 * las fotos en una carpeta de Drive y hace el análisis con la API de Gemini.
 *
 * Propiedades del script (Configuración del proyecto → Propiedades del script):
 *   CLAVE_ACCESO          clave que se pide al abrir la app (obligatoria)
 *   GEMINI_API_KEY        clave gratuita de Google AI Studio (obligatoria para el análisis)
 *   GEMINI_MODEL          opcional, por omisión gemini-3.8-flash
 *   GEMINI_MODEL_RESPALDO opcional, por omisión gemini-3.5-flash-lite (se usa si se agota la cuota del principal)
 * La hoja y la carpeta se crean solas la primera vez (HOJA_ID y CARPETA_ID).
 */
const VERSION = '1.0.0';
const COLS = ['establecimientos', 'simulacros'];
const TROZO = 45000;           // caracteres por celda (límite de Sheets: 50,000)
const MAX_VIDEO = 45 * 1024 * 1024;

function doGet() {
  return salida({ ok: true, app: 'Evaluación de Simulacros VIME', version: VERSION });
}

function doPost(e) {
  let r;
  try {
    const req = JSON.parse(e.postData.contents);
    const P = PropertiesService.getScriptProperties();
    const clave = P.getProperty('CLAVE_ACCESO');
    if (!clave) throw new Error('Falta configurar CLAVE_ACCESO en las propiedades del script.');
    if (String(req.clave || '') !== clave) {
      Utilities.sleep(1000);
      return salida({ ok: false, error: 'clave', mensaje: 'Clave de acceso incorrecta.' });
    }
    const f = ACCIONES[req.accion];
    if (!f) throw new Error('Acción desconocida: ' + req.accion);
    r = Object.assign({ ok: true }, f(req, P));
  } catch (err) {
    r = { ok: false, error: err.codigo || 'error', mensaje: String(err.message || err) };
  }
  return salida(r);
}

function salida(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function fallo(codigo, mensaje) { const e = new Error(mensaje); e.codigo = codigo; return e; }

/* ---------------- Almacenamiento ---------------- */
function libro(P) {
  let id = P.getProperty('HOJA_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { /* se recrea */ } }
  const ss = SpreadsheetApp.create('Evaluación de Simulacros VIME · Datos');
  COLS.forEach((c, i) => {
    const sh = i === 0 ? ss.getSheets()[0].setName(c) : ss.insertSheet(c);
    sh.getRange(1, 1, 1, 3).setValues([['id', 'actualizado', 'datos (JSON)']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  P.setProperty('HOJA_ID', ss.getId());
  return ss;
}
function carpeta(P) {
  let id = P.getProperty('CARPETA_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* se recrea */ } }
  const f = DriveApp.createFolder('Evaluación de Simulacros VIME · Fotos');
  P.setProperty('CARPETA_ID', f.getId());
  return f;
}
function hoja(P, col) {
  if (COLS.indexOf(col) < 0) throw new Error('Colección inválida');
  return libro(P).getSheetByName(col);
}
function leerTodo(sh) {
  const n = sh.getLastRow();
  if (n < 2) return [];
  const w = sh.getLastColumn();
  return sh.getRange(2, 1, n - 1, w).getValues().filter(r => r[0]).map(r => {
    const txt = r.slice(2).join('');
    try { return JSON.parse(txt); } catch (e) { return null; }
  }).filter(Boolean);
}
function filaDe(sh, id) {
  const t = sh.getRange('A:A').createTextFinder(String(id)).matchEntireCell(true).findNext();
  return t ? t.getRow() : 0;
}

const ACCIONES = {
  ping(req, P) {
    libro(P); carpeta(P);
    return { version: VERSION, ia: !!P.getProperty('GEMINI_API_KEY'), modelo: P.getProperty('GEMINI_MODEL') || 'gemini-3.8-flash' };
  },
  listar(req, P) {
    const out = {};
    COLS.forEach(c => out[c] = leerTodo(hoja(P, c)));
    return out;
  },
  guardar(req, P) {
    const sh = hoja(P, req.col);
    const obj = req.obj || {};
    obj.id = String(req.id);
    const txt = JSON.stringify(obj);
    const partes = [];
    for (let i = 0; i < txt.length; i += TROZO) partes.push(txt.slice(i, i + TROZO));
    const fila = [obj.id, new Date()].concat(partes);
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      let r = filaDe(sh, obj.id);
      if (!r) r = sh.getLastRow() + 1;
      const w = Math.max(sh.getLastColumn(), fila.length);
      const vacia = new Array(w).fill('');
      fila.forEach((v, i) => vacia[i] = v);
      sh.getRange(r, 1, 1, w).setValues([vacia]);
    } finally { lock.releaseLock(); }
    return {};
  },
  borrar(req, P) {
    const sh = hoja(P, req.col);
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try { const r = filaDe(sh, req.id); if (r) sh.deleteRow(r); } finally { lock.releaseLock(); }
    return {};
  },
  subirFoto(req, P) {
    const bytes = Utilities.base64Decode(req.b64);
    const blob = Utilities.newBlob(bytes, req.mime || 'image/jpeg', req.nombre || ('foto_' + Date.now() + '.jpg'));
    const raiz = carpeta(P);
    const nom = 'Simulacro ' + (req.simId || 'sin-id');
    const it = raiz.getFoldersByName(nom);
    const sub = it.hasNext() ? it.next() : raiz.createFolder(nom);
    const f = sub.createFile(blob);
    return { fileId: f.getId() };
  },
  foto(req, P) {
    const b = DriveApp.getFileById(req.fileId).getBlob();
    return { b64: Utilities.base64Encode(b.getBytes()), mime: b.getContentType() };
  },
  borrarFoto(req, P) {
    try { DriveApp.getFileById(req.fileId).setTrashed(true); } catch (e) { }
    return {};
  },
  /** Revisa que el enlace de Drive se pueda leer y sube el video a Gemini (dura 48 h allá). */
  subirVideo(req, P) {
    const id = idDrive(req.url);
    if (!id) throw fallo('video', 'El enlace no parece de Google Drive.');
    let file;
    try { file = DriveApp.getFileById(id); } catch (e) { throw fallo('video', 'No se puede abrir el video con su cuenta de Google. Verifique que el archivo sea suyo o esté compartido con usted.'); }
    const size = file.getSize();
    if (size > MAX_VIDEO) throw fallo('video', 'El video pesa ' + Math.round(size / 1048576) + ' MB; el máximo para análisis directo es 45 MB. Use «Extraer cuadros» o recorte el video.');
    const blob = file.getBlob();
    const mime = blob.getContentType() || 'video/mp4';
    if (!/^video\//.test(mime)) throw fallo('video', 'El archivo de Drive no es un video (' + mime + ').');
    const key = claveIA(P);
    const ini = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/upload/v1beta/files', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-goog-api-key': key, 'X-Goog-Upload-Protocol': 'resumable', 'X-Goog-Upload-Command': 'start',
        'X-Goog-Upload-Header-Content-Length': String(size), 'X-Goog-Upload-Header-Content-Type': mime },
      payload: JSON.stringify({ file: { display_name: file.getName() } })
    });
    const hs = ini.getAllHeaders(); let up = '';
    Object.keys(hs).forEach(k => { if (k.toLowerCase() === 'x-goog-upload-url') up = hs[k]; });
    if (!up) throw fallo('ia', 'Gemini no aceptó la carga del video: ' + ini.getContentText().slice(0, 300));
    const r = UrlFetchApp.fetch(up, { method: 'post', muteHttpExceptions: true, contentType: mime,
      headers: { 'X-Goog-Upload-Offset': '0', 'X-Goog-Upload-Command': 'upload, finalize' }, payload: blob.getBytes() });
    const info = JSON.parse(r.getContentText()).file;
    if (!info) throw fallo('ia', 'No se completó la carga del video: ' + r.getContentText().slice(0, 300));
    let st = info, t0 = Date.now();
    while (st.state === 'PROCESSING' && Date.now() - t0 < 240000) {
      Utilities.sleep(4000);
      st = JSON.parse(UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/' + info.name, { headers: { 'x-goog-api-key': key }, muteHttpExceptions: true }).getContentText());
    }
    if (st.state !== 'ACTIVE') throw fallo('ia', 'Gemini no terminó de procesar el video (' + st.state + '). Intente de nuevo en unos minutos.');
    return { uri: st.uri, mime: st.mimeType || mime, nombre: file.getName(), expira: st.expirationTime || '' };
  },
  /** Una consulta a Gemini con texto, fotos de Drive y videos ya subidos. */
  ia(req, P) {
    const parts = [{ text: String(req.prompt || '') }];
    (req.fotos || []).forEach((fid, i) => {
      const b = DriveApp.getFileById(fid).getBlob();
      parts.push({ text: 'Foto ' + (i + 1 + (req.desde || 0)) + ':' });
      parts.push({ inline_data: { mime_type: b.getContentType() || 'image/jpeg', data: Utilities.base64Encode(b.getBytes()) } });
    });
    (req.videos || []).forEach(v => {
      parts.push({ text: 'Video ' + v.etiqueta + ':' });
      parts.push({ file_data: { mime_type: v.mime, file_uri: v.uri } });
    });
    return gemini(P, parts, req.json !== false);
  }
};

function claveIA(P) {
  const k = P.getProperty('GEMINI_API_KEY');
  if (!k) throw fallo('ia', 'Falta configurar GEMINI_API_KEY en las propiedades del script.');
  return k;
}
function idDrive(url) {
  const s = String(url || '');
  const m = s.match(/\/d\/([\w-]{20,})/) || s.match(/[?&]id=([\w-]{20,})/);
  return m ? m[1] : (/^[\w-]{20,}$/.test(s) ? s : '');
}
function gemini(P, parts, json) {
  const key = claveIA(P);
  const modelos = [P.getProperty('GEMINI_MODEL') || 'gemini-3.8-flash', P.getProperty('GEMINI_MODEL_RESPALDO') || 'gemini-3.5-flash-lite'];
  const body = { contents: [{ role: 'user', parts: parts }], generationConfig: { temperature: 0.2, maxOutputTokens: 32768 } };
  if (json) body.generationConfig.responseMimeType = 'application/json';
  let ultimo = '';
  for (let i = 0; i < modelos.length; i++) {
    const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + modelos[i] + ':generateContent', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-goog-api-key': key }, payload: JSON.stringify(body)
    });
    const code = res.getResponseCode(); const txt = res.getContentText();
    if (code === 200) {
      const d = JSON.parse(txt);
      const c = d.candidates && d.candidates[0];
      if (!c || !c.content) throw fallo('ia', 'Gemini no devolvió respuesta (' + ((c && c.finishReason) || (d.promptFeedback && d.promptFeedback.blockReason) || 'sin motivo') + ').');
      const texto = c.content.parts.filter(p => !p.thought).map(p => p.text || '').join('');
      return { texto: texto, modelo: modelos[i], recortado: c.finishReason === 'MAX_TOKENS' };
    }
    ultimo = code + ' ' + txt.slice(0, 400);
    if (code === 429 || code === 503 || code === 404) continue;  // cuota agotada, saturado o modelo retirado → respaldo
    throw fallo('ia', 'Error de Gemini: ' + ultimo);
  }
  throw fallo('cuota', 'Se agotó por hoy la cuota gratuita de Gemini (o el servicio está saturado). Intente más tarde. Detalle: ' + ultimo);
}

/** Ejecute esta función una vez desde el editor para autorizar permisos y comprobar la configuración. */
function instalar() {
  const P = PropertiesService.getScriptProperties();
  const ss = libro(P), f = carpeta(P);
  Logger.log('Hoja de datos: ' + ss.getUrl());
  Logger.log('Carpeta de fotos: ' + f.getUrl());
  Logger.log('CLAVE_ACCESO: ' + (P.getProperty('CLAVE_ACCESO') ? 'configurada' : 'FALTA'));
  if (P.getProperty('GEMINI_API_KEY')) {
    const r = gemini(P, [{ text: 'Responde solo con la palabra: listo' }], false);
    Logger.log('Gemini (' + r.modelo + '): ' + r.texto);
  } else Logger.log('GEMINI_API_KEY: FALTA');
}
