// Consulta de listas públicas para el KYC de Certus (OFAC, ONU, SAT 69 y 69-B).
// Las listas se descargan y se guardan en memoria; se renuevan cada 24 horas.
// La Lista de Personas Bloqueadas de la UIF y la Consulta PEP 2.0 no son públicas:
// las consulta el propio sujeto obligado con su e.firma, por eso no están aquí.
const axios = require('axios');

const RFC_RE = /^[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}$/;
const VIGENCIA_MS = 24 * 60 * 60 * 1000;
const STOP = new Set(['DE', 'DEL', 'LA', 'LAS', 'LOS', 'EL', 'Y', 'E', 'THE', 'OF', 'AND', 'SA', 'CV', 'S', 'A', 'C', 'V', 'SC', 'SAPI', 'SAB', 'RL', 'SRL', 'SAS', 'AC', 'LLC', 'INC', 'LTD', 'CO', 'CORP', 'BIN', 'AL', 'ABU']);

const FUENTES = {
  ofac: {
    nombre: 'OFAC (SDN y lista consolidada, EE. UU.)',
    urls: [
      'https://www.treasury.gov/ofac/downloads/sdn.csv',
      'https://www.treasury.gov/ofac/downloads/alt.csv',
      'https://www.treasury.gov/ofac/downloads/consolidated/cons_prim.csv',
      'https://www.treasury.gov/ofac/downloads/consolidated/cons_alt.csv'
    ],
    tipo: 'nombres'
  },
  onu: {
    nombre: 'Consejo de Seguridad de la ONU (lista consolidada)',
    urls: ['https://scsanctions.un.org/resources/xml/en/consolidated.xml'],
    tipo: 'nombres'
  },
  sat69b: {
    nombre: 'SAT 69-B (operaciones inexistentes)',
    urls: ['http://omawww.sat.gob.mx/cifras_sat/Documents/Listado_Completo_69-B.csv'],
    tipo: 'rfc'
  },
  sat69: {
    nombre: 'SAT 69 (firmes, exigibles, no localizados, sentencias)',
    urls: [
      'http://omawww.sat.gob.mx/cifras_sat/Documents/Firmes.csv',
      'http://omawww.sat.gob.mx/cifras_sat/Documents/Exigibles.csv',
      'http://omawww.sat.gob.mx/cifras_sat/Documents/No%20localizados.csv',
      'http://omawww.sat.gob.mx/cifras_sat/Documents/Sentencias.csv'
    ],
    tipo: 'rfc'
  }
};

const estado = {};      // id -> {ok, cargada, registros, error, cargando}
const datos = {};       // id -> {nombres: [{n, tokens, detalle}], rfc: Map}

function normaliza(t) {
  return String(t || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/Ñ/g, 'N').replace(/[^A-Z0-9 ]+/g, ' ').replace(/\s+/g, ' ').trim();
}
function tokens(t) { return normaliza(t).split(' ').filter(w => w.length > 1 && !STOP.has(w)); }

function decodifica(buf) {
  try { return new TextDecoder('utf-8', { fatal: true }).decode(buf); }
  catch (e) { return Buffer.from(buf).toString('latin1'); }
}

// CSV con comillas
function filasCSV(texto, cb) {
  let campo = '', fila = [], q = false;
  for (let i = 0; i < texto.length; i++) {
    const c = texto[i];
    if (q) {
      if (c === '"') { if (texto[i + 1] === '"') { campo += '"'; i++; } else q = false; }
      else campo += c;
    } else if (c === '"') q = true;
    else if (c === ',') { fila.push(campo); campo = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && texto[i + 1] === '\n') i++;
      fila.push(campo); campo = '';
      if (fila.length > 1 || fila[0]) cb(fila);
      fila = [];
    } else campo += c;
  }
  if (campo || fila.length) { fila.push(campo); cb(fila); }
}

async function bajar(url) {
  const r = await axios.get(url, {
    responseType: 'arraybuffer', timeout: 180000, maxContentLength: Infinity, maxBodyLength: Infinity,
    headers: { 'User-Agent': 'Mozilla/5.0 (Certus PLD; consulta de listas)' }
  });
  return decodifica(r.data);
}

const parsers = {
  async ofac(textos) {
    const L = [];
    const prim = new Map();
    // sdn.csv / cons_prim.csv: ent_num, nombre, tipo, programa, ...
    [textos[0], textos[2]].forEach((t, k) => {
      if (!t) return;
      filasCSV(t, f => {
        if (f.length < 4 || !/^\d+$/.test(f[0].trim())) return;
        const n = f[1].trim(); if (!n || n === '-0-') return;
        const det = (k ? 'Consolidada' : 'SDN') + ' · ' + (f[2] && f[2] !== '-0-' ? f[2].trim() + ' · ' : '') + 'Programa ' + f[3].trim();
        prim.set((k ? 'c' : 's') + f[0].trim(), { n, det });
        L.push({ n, det });
      });
    });
    // alt.csv / cons_alt.csv: ent_num, alt_num, tipo, nombre alterno
    [textos[1], textos[3]].forEach((t, k) => {
      if (!t) return;
      filasCSV(t, f => {
        if (f.length < 4 || !/^\d+$/.test(f[0].trim())) return;
        const n = f[3].trim(); if (!n || n === '-0-') return;
        const p = prim.get((k ? 'c' : 's') + f[0].trim());
        L.push({ n, det: 'Alias de ' + (p ? p.n : 'registro ' + f[0].trim()) + (p ? ' · ' + p.det : '') });
      });
    });
    return { nombres: L };
  },
  async onu(textos) {
    const x = textos[0] || '', L = [];
    const tag = (b, t) => { const m = b.match(new RegExp('<' + t + '>([\\s\\S]*?)</' + t + '>')); return m ? m[1].trim() : ''; };
    const dec = s => s.replace(/&amp;/g, '&').replace(/&apos;/g, "'").replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>');
    for (const [bloque, tipo] of [['INDIVIDUAL', 'Persona'], ['ENTITY', 'Entidad']]) {
      const re = new RegExp('<' + bloque + '>([\\s\\S]*?)</' + bloque + '>', 'g'); let m;
      while ((m = re.exec(x))) {
        const b = m[1];
        const n = dec(['FIRST_NAME', 'SECOND_NAME', 'THIRD_NAME', 'FOURTH_NAME'].map(t => tag(b, t)).filter(Boolean).join(' '));
        const ref = tag(b, 'REFERENCE_NUMBER'), lista = tag(b, 'UN_LIST_TYPE');
        const det = tipo + ' · ' + (lista ? 'Comité ' + lista : '') + (ref ? ' · Ref. ' + ref : '');
        if (n) L.push({ n, det });
        const ra = new RegExp('<' + bloque + '_ALIAS>([\\s\\S]*?)</' + bloque + '_ALIAS>', 'g'); let a;
        while ((a = ra.exec(b))) { const an = dec(tag(a[1], 'ALIAS_NAME')); if (an) L.push({ n: an, det: 'Alias de ' + n + ' · ' + det }); }
      }
    }
    return { nombres: L };
  },
  async sat69b(textos) {
    const rfc = new Map();
    filasCSV(textos[0] || '', f => {
      const i = f.findIndex(c => RFC_RE.test(c.trim().toUpperCase()));
      if (i < 0) return;
      const sit = f.slice(i + 2).map(c => c.trim()).find(c => /definitivo|presunto|desvirtuado|favorable/i.test(c)) || '';
      rfc.set(f[i].trim().toUpperCase(), { n: (f[i + 1] || '').trim(), det: 'Situación: ' + (sit || 'ver listado') });
    });
    return { rfc };
  },
  async sat69(textos) {
    const nombresLista = ['Créditos firmes', 'Créditos exigibles', 'No localizados', 'Sentencias'];
    const rfc = new Map();
    textos.forEach((t, k) => {
      if (!t) return;
      filasCSV(t, f => {
        const i = f.findIndex(c => RFC_RE.test(c.trim().toUpperCase()));
        if (i < 0) return;
        const r = f[i].trim().toUpperCase(); const prev = rfc.get(r);
        if (prev) { if (!prev.det.includes(nombresLista[k])) prev.det += ', ' + nombresLista[k]; }
        else rfc.set(r, { n: (f[i + 1] || '').trim().slice(0, 120), det: 'Supuesto: ' + nombresLista[k] });
      });
    });
    return { rfc };
  }
};

async function cargar(id) {
  const F = FUENTES[id]; const e = estado[id] || (estado[id] = {});
  if (e.cargando) return e.cargando;
  e.cargando = (async () => {
    const textos = []; const errores = [];
    for (const u of F.urls) {
      try { textos.push(await bajar(u)); }
      catch (err) { textos.push(null); errores.push(u.split('/').pop() + ': ' + (err.response ? 'HTTP ' + err.response.status : err.code || err.message)); }
    }
    if (textos.every(t => !t)) throw new Error(errores.join('; '));
    const d = await parsers[id](textos);
    if (d.nombres) d.nombres.forEach(x => { x.t = tokens(x.n); });
    const registros = d.nombres ? d.nombres.length : d.rfc.size;
    if (!registros) throw new Error('El archivo no tiene el formato esperado' + (errores.length ? ' (' + errores.join('; ') + ')' : ''));
    datos[id] = d;
    Object.assign(e, { ok: true, cargada: new Date().toISOString(), registros, error: errores.length ? 'Parcial: ' + errores.join('; ') : '' });
  })().catch(err => {
    Object.assign(e, { ok: !!datos[id], error: String(err.message || err).slice(0, 300) });
    if (!e.cargada) e.cargada = null;
  }).finally(() => { e.cargando = null; });
  return e.cargando;
}

async function asegurar() {
  const ahora = Date.now();
  await Promise.all(Object.keys(FUENTES).map(id => {
    const e = estado[id];
    if (!e || !e.cargada || ahora - Date.parse(e.cargada) > VIGENCIA_MS) return cargar(id);
    return null;
  }));
}

function buscarNombre(id, consulta) {
  const q = tokens(consulta); if (q.length < 2) return [];
  const qs = new Set(q); const out = [];
  for (const x of datos[id].nombres) {
    if (x.t.length < 2) continue;
    const ts = new Set(x.t);
    const todosQ = q.every(w => ts.has(w));
    const todosX = x.t.every(w => qs.has(w));
    if (todosQ || todosX) {
      const comunes = q.filter(w => ts.has(w)).length;
      if (comunes < 2) continue;
      out.push({ encontrado: x.n, detalle: x.det, exacto: todosQ && todosX });
      if (out.length >= 15) break;
    }
  }
  return out;
}

async function consultar({ rfc, nombres }) {
  await asegurar();
  const r = String(rfc || '').toUpperCase().replace(/[^A-Z0-9Ñ&]/g, '');
  const ns = (Array.isArray(nombres) ? nombres : String(nombres || '').split('|')).map(s => String(s).trim()).filter(Boolean).slice(0, 10);
  const resultados = [];
  for (const id of Object.keys(FUENTES)) {
    if (!datos[id]) continue;
    if (FUENTES[id].tipo === 'rfc') {
      if (r && datos[id].rfc.has(r)) { const x = datos[id].rfc.get(r); resultados.push({ fuente: id, por: 'RFC', consulta: r, encontrado: x.n, detalle: x.det, exacto: true }); }
    } else {
      ns.forEach(n => buscarNombre(id, n).forEach(h => resultados.push(Object.assign({ fuente: id, por: 'nombre', consulta: n }, h))));
    }
  }
  const fuentes = Object.keys(FUENTES).map(id => ({ id, nombre: FUENTES[id].nombre, ok: !!datos[id], cargada: (estado[id] || {}).cargada || null, registros: (estado[id] || {}).registros || 0, error: (estado[id] || {}).error || '' }));
  return { fecha: new Date().toISOString(), rfc: r, nombres: ns, fuentes, resultados };
}

module.exports = { consultar, asegurar, estado, FUENTES, _test: { tokens, buscarNombre, datos, parsers, filasCSV } };
