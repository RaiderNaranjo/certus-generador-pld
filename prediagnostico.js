// ─────────────────────────────────────────────
// PRE-DIAGNÓSTICO PLD/FT — cálculo y PDF para el cliente
// Entrada: { cliente, perfil, respuestas }  (lo arma el intake)
// Salida:  { resultado, buffer }  (PDF carta, 3 páginas)
// Fundamento verificado contra LFPIORPI, última reforma DOF 16-07-2025.
// ─────────────────────────────────────────────
const fs = require('fs');
const path = require('path');
const PdfPrinter = require('pdfmake');

const NAVY = '#1F3864', GOLD = '#B0763A', INK = '#1B1F2A', MUTED = '#6B7280', LINE = '#E6E1D8', SOFT = '#F7F3EE';
const COLOR = { r: '#B3322B', a: '#D9822B', v: '#3F8F5A', g: '#9AA0AB' };

const printer = new PdfPrinter({
  Helvetica: { normal: 'Helvetica', bold: 'Helvetica-Bold', italics: 'Helvetica-Oblique', bolditalics: 'Helvetica-BoldOblique' },
  Times: { normal: 'Times-Roman', bold: 'Times-Bold', italics: 'Times-Italic', bolditalics: 'Times-BoldItalic' }
});

const asset = f => { try { return 'data:image/png;base64,' + fs.readFileSync(path.join(__dirname, 'assets', f)).toString('base64'); } catch (e) { return null; } };
const LOGO_C = asset('alianza_certus.png'), LOGO_E = asset('alianza_es.png');

// ── Modelo ──
const VAL = { si: 0, par: 0.5, no: 1, ns: 0.75 };
const BLOQUES = {
  reg: ['Alta y registro en el padrón', 'Art. 18, fr. IV Bis'],
  rep: ['Representante Encargado de Cumplimiento', 'Art. 20'],
  man: ['Manual y evaluación de riesgos', 'Art. 18, fr. VII y VIII'],
  kyc: ['Identificación de clientes', 'Art. 18, fr. I a IV'],
  avi: ['Avisos ante la SHCP', 'Arts. 17, 18 fr. VI y 23'],
  cap: ['Capacitación', 'Arts. 18 fr. IX y 20'],
  aud: ['Monitoreo y auditoría', 'Art. 18, fr. X y XI']
};
// [id, bloque, peso, tema para la sesión]
const PREGUNTAS = [
  ['reg', 'reg', 3, 'Confirmar el estatus de alta y registro en el SPPLD'],
  ['rep', 'rep', 3, 'Confirmar la designación y vigencia del Representante'],
  ['man', 'man', 3, 'Revisar si existe Manual de Políticas Internas y su contenido'],
  ['ebr', 'man', 2, 'Revisar si existe evaluación con enfoque basado en riesgo'],
  ['kyc', 'kyc', 3, 'Revisar una muestra de expedientes de clientes'],
  ['bc', 'kyc', 2, 'Revisar cómo se identifica al beneficiario controlador y a las PEP'],
  ['cons', 'kyc', 1, 'Confirmar dónde y cómo se conserva la documentación'],
  ['avi', 'avi', 3, 'Revisar el historial de Avisos presentados y sus acuses'],
  ['a24', 'avi', 2, 'Revisar el procedimiento ante operaciones sospechosas'],
  ['cap', 'cap', 2, 'Confirmar la capacitación del año y sus constancias'],
  ['mon', 'aud', 2, 'Revisar el mecanismo de monitoreo y acumulación'],
  ['aud', 'aud', 1, 'Confirmar si existe revisión o auditoría del último año']
];
const NIVELES = ['Contenida (preliminar)', 'Media', 'Alta', 'Crítica'];

const norm = v => (VAL[v] !== undefined ? v : 'ns');

function calcular({ cliente = {}, perfil = {}, respuestas = {} }) {
  const pf = String(cliente.TipoPersona || '').toUpperCase() === 'PF';
  const R = {};
  PREGUNTAS.forEach(([id]) => { R[id] = norm(respuestas[id]); });
  const activas = PREGUNTAS.filter(q => !(pf && q[0] === 'rep'));

  let sw = 0, sv = 0;
  activas.forEach(([id, , w]) => { sw += w; sv += w * VAL[R[id]]; });
  const indice = Math.round((100 * sv) / sw);
  const noSe = activas.filter(q => R[q[0]] === 'ns').length;

  // Riesgo inherente: volumen + efectivo + clientes extranjeros/estructuras complejas
  const vol = { '1-5': 0, '6-20': 1, '21-50': 2, '51+': 2 }[perfil.operacionesMes] || 0;
  const efe = Math.min(2, Math.max(0, parseInt(perfil.efectivo, 10) || 0));
  const ext = perfil.extranjeros === 'no' ? 0 : 1;
  const inh = vol + efe + ext;
  const inherente = inh >= 3 ? 'Elevado' : inh >= 1 ? 'Medio' : 'Menor';

  let li = indice >= 60 ? 3 : indice >= 40 ? 2 : indice >= 20 ? 1 : 0;
  if (inh >= 3 && li > 0 && li < 3) li++;
  const indeterminado = noSe >= 4;
  const nivel = indeterminado ? 'No determinada' : NIVELES[li];
  const color = indeterminado ? 'g' : ['v', 'a', 'a', 'r'][li];

  const bloques = Object.keys(BLOQUES).map(b => {
    if (pf && b === 'rep') return { b, na: true };
    const qs = PREGUNTAS.filter(q => q[1] === b);
    let w = 0, v = 0, todosNs = true;
    qs.forEach(([id, , qw]) => { w += qw; v += qw * VAL[R[id]]; if (R[id] !== 'ns') todosNs = false; });
    const x = v / w;
    if (todosNs) return { b, c: 'g', t: 'Por confirmar' };
    return x >= 0.6 ? { b, c: 'r', t: 'Brecha relevante' } : x >= 0.25 ? { b, c: 'a', t: 'Requiere atención' } : { b, c: 'v', t: 'Sin brecha aparente' };
  });

  const pendientes = activas
    .filter(([id]) => R[id] === 'ns' || R[id] === 'par')
    .map(([id, , , tema]) => tema + (R[id] === 'ns' ? ' (respondió «No sé»)' : ' (respondió «Parcialmente»)'));
  if (perfil.extranjeros === 'ns') pendientes.push('Precisar si hay clientes extranjeros o estructuras societarias complejas');
  if (perfil.verificacion === 'ns') pendientes.push('Confirmar si existe algún requerimiento o visita de la autoridad en curso');
  pendientes.push('Validar las respuestas con la documentación soporte', 'Revisar el número y monto de operaciones de los últimos 12 meses');

  return { pf, respuestas: R, indice, nivel, color, inherente, noSe, bloques, pendientes };
}

// ── Helpers de maquetación ──
const UMA = () => parseFloat(process.env.UMA_VALOR || '117.31');
const mx = n => '$' + Math.round(n).toLocaleString('en-US');
const punto = c => ({ canvas: [{ type: 'ellipse', x: 5, y: 6.5, r1: 4.2, r2: 4.2, color: COLOR[c] }], width: 13 });
const chip = (c, t) => ({ columns: [punto(c), { text: t, bold: true, fontSize: 9.5 }], columnGap: 2 });
const h1 = t => ({ text: t, font: 'Times', bold: true, fontSize: 20, color: NAVY, margin: [0, 0, 0, 8] });
const h2 = t => ({ text: t, font: 'Times', bold: true, fontSize: 13.5, color: NAVY, margin: [0, 14, 0, 6] });
const kick = t => ({ text: t, bold: true, fontSize: 8.5, color: GOLD, margin: [0, 0, 0, 4] });
const caja = (contenido, alerta) => ({
  table: { widths: ['*'], body: [[{ stack: [].concat(contenido), fontSize: 9.5, margin: [4, 4, 4, 4] }]] },
  layout: { hLineColor: () => (alerta ? '#E7C79F' : LINE), vLineColor: () => (alerta ? '#E7C79F' : LINE), fillColor: () => (alerta ? '#FDF6EC' : '#FBF9F5') },
  margin: [0, 6, 0, 6]
});
const tabla = (widths, head, rows) => ({
  table: { headerRows: 1, widths, body: [head.map(h => ({ text: h, bold: true, fontSize: 8.5, color: NAVY })), ...rows] },
  layout: {
    hLineWidth: (i) => (i === 1 ? 1.5 : 0.6), vLineWidth: () => 0,
    hLineColor: (i) => (i === 1 ? NAVY : LINE), paddingTop: () => 5, paddingBottom: () => 5
  },
  fontSize: 9.5
});


// ── Aviso de privacidad: constancia de aceptación ──
const fechaMx = iso => { const d = iso ? new Date(iso) : new Date(); return isNaN(d) ? '' : d.toLocaleString('es-MX', { day: 'numeric', month: 'long', year: 'numeric', hour: '2-digit', minute: '2-digit', timeZone: 'America/Mexico_City' }); };
function constanciaPrivacidad(pv = {}) {
  if (pv.acepta !== 'Sí') return '';
  return `Usted aceptó el aviso de privacidad${pv.version ? ' (versión ' + pv.version + ')' : ''} el ${fechaMx(pv.fecha)} al enviar el formulario${pv.secundarias === 'No' ? ' y se opuso al uso de sus datos para finalidades secundarias' : ''}.`;
}

// ── Correo al cliente: acuse + diagnóstico preliminar + invitación a agendar ──
const esc = t => String(t == null ? '' : t).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
function correoCliente(datos, r) {
  const c = datos.cliente || {}, pv = datos.privacidad || {};
  const agenda = process.env.AGENDA_URL || '';
  const nombre = esc(c.NombreRepresentanteLegal || 'cliente');
  const colores = { r: '#B3322B', a: '#D9822B', v: '#3F8F5A', g: '#6B7280' };
  const boton = agenda
    ? `<p style="text-align:center;margin:22px 0;"><a href="${esc(agenda)}" style="background:#1F3864;color:#ffffff;text-decoration:none;padding:12px 26px;font-weight:bold;border-radius:4px;display:inline-block;">Agendar mi sesión de diagnóstico</a></p>`
    : `<p style="background:#1F3864;color:#ffffff;padding:12px 16px;"><strong style="color:#F1C38C;">Para agendar:</strong> responda este correo con dos o tres horarios que le acomoden, o llámenos al 999 163 1363.</p>`;
  const asunto = `Recibimos su información — Diagnóstico preliminar PLD/FT de ${c.La_Empresa || 'su empresa'}`;
  const html = `<div style="font-family:Arial,Helvetica,sans-serif;max-width:600px;margin:0 auto;color:#1F3864;">
<div style="background:#1F3864;padding:18px 24px;border-bottom:3px solid #C9A84C;"><span style="color:#C9A84C;font-size:20px;letter-spacing:2px;font-weight:bold;">CERTUS</span> <span style="color:#ffffff;font-size:12px;letter-spacing:2px;">CONSULTORES</span><span style="color:#C9A84C;font-size:16px;padding:0 12px;">|</span><span style="color:#C9A84C;font-size:16px;letter-spacing:2px;font-weight:bold;">ESPINOSA-SALAZAR</span> <span style="color:#ffffff;font-size:12px;letter-spacing:2px;">ABOGADOS CORPORATIVOS</span></div>
<div style="padding:24px;font-size:15px;line-height:1.6;color:#1B1F2A;">
<p>Estimado/a <strong>${nombre}</strong>:</p>
<p>Recibimos correctamente la información de <strong>${esc(c.La_Empresa)}</strong> que nos compartió en nuestro formulario de registro. Gracias por su confianza.</p>
<p>Con sus respuestas elaboramos un <strong>diagnóstico preliminar de exposición en materia PLD/FT</strong>, que encontrará adjunto en PDF (folio ${esc(r.folio)}).</p>
<table role="presentation" style="width:100%;border-collapse:collapse;background:#F7F3EE;border-left:4px solid #D4944A;margin:8px 0 16px;"><tr>
<td style="padding:14px 16px;font-size:34px;font-weight:bold;color:#1F3864;white-space:nowrap;">${r.indice}<span style="font-size:14px;color:#6B7280;"> /100</span></td>
<td style="padding:14px 16px 14px 0;font-size:14px;">Exposición preliminar: <strong style="color:${colores[r.color]};">${esc(r.nivel)}</strong><br>Riesgo inherente de la actividad: <strong>${esc(r.inherente)}</strong></td></tr></table>
<p>Este resultado es una primera lectura basada solo en lo que usted declaró. <strong>El diagnóstico completo se entrega en una sesión personal</strong>, presencial o por videollamada, de unos 45 minutos. En ella revisamos sus documentos, resolvemos las preguntas que quedaron pendientes y le presentamos la ruta para regularizarse.</p>
<p><strong>Para continuar, es necesario que agende su sesión.</strong></p>
${boton}
<p style="font-size:14px;">Si los tiene a la mano, le pedimos preparar: acuse de alta en el SPPLD, designación del Representante de Cumplimiento, manual vigente, acuses de Avisos de los últimos 12 meses y dos expedientes de clientes.</p>
<p>Si desea agregar información o tiene alguna duda, puede responder a este correo o escribirnos a <a href="mailto:contacto@certusconsultores.com.mx">contacto@certusconsultores.com.mx</a> o a <a href="mailto:contacto@espinosasalazarabogados.com">contacto@espinosasalazarabogados.com</a>.</p>
<p>Atentamente,</p>
<p><strong>Certus Consultores</strong><br>Cumplimiento normativo PLD/FT<br>Mérida, Yucatán · Tel. 999 163 1363</p>
<p><strong>Espinosa-Salazar Abogados Corporativos</strong><br>Tel. 999 575 3129</p>
</div>
<div style="background:#f3f1ed;padding:12px 24px;font-size:11px;color:#5a5650;line-height:1.5;">
<p style="margin:0 0 6px;"><strong>Aviso de privacidad.</strong> ${esc(constanciaPrivacidad(pv)) || 'Su información se trata conforme a nuestro aviso de privacidad.'} Responsable: José Luis Naranjo Bobadilla (Certus Consultores). Consulte el aviso integral en <a href="https://certusconsultores.com.mx/aviso-de-privacidad/">certusconsultores.com.mx/aviso-de-privacidad</a>. Puede ejercer sus derechos ARCO o revocar su consentimiento en contacto@certusconsultores.com.mx.</p>
<p style="margin:0;">El diagnóstico adjunto es orientativo y se generó automáticamente a partir de la información que usted declaró; no constituye dictamen ni opinión jurídica.</p>
</div></div>`;
  return { asunto, html };
}

function generarPDF(datos) {
  const { cliente = {}, perfil = {} } = datos;
  const r = calcular(datos);
  const uma = UMA();
  const hoy = new Date();
  const fecha = hoy.toLocaleDateString('es-MX', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'America/Mexico_City' });
  const folio = 'PRE-' + hoy.toLocaleDateString('en-CA', { timeZone: 'America/Mexico_City' }).replace(/-/g, '') + '-' + String(cliente.RFC || 'SINRFC').toUpperCase().slice(0, 13);
  const agenda = process.env.AGENDA_URL || '';
  const tipo = String(cliente.Tipo || '');
  const fedatario = tipo === 'notarial' || tipo === 'corredores';

  // Supuestos de sanción según brechas
  const R = r.respuestas;
  const brecha18 = PREGUNTAS.some(([id, b]) => b !== 'avi' && b !== 'rep' && R[id] !== 'si');
  const brechaAv = R.avi !== 'si';
  const efectivo = String(perfil.efectivo || '0') !== '0';
  const rango = (a, b) => `${mx(a * uma)} a ${mx(b * uma)}`;
  const sanciones = [];
  if (brecha18) sanciones.push(['Incumplir obligaciones del artículo 18 (registro, manual, identificación, capacitación, monitoreo, auditoría)', 'Arts. 53 fr. II y 54 fr. I', '200 a 2,000 UMA', rango(200, 2000)]);
  if (brechaAv) {
    sanciones.push(['Presentar Avisos fuera de plazo (hasta 30 días de retraso)', 'Arts. 53 fr. III y 54 fr. I', '200 a 2,000 UMA', rango(200, 2000)]);
    sanciones.push(['Omitir Avisos (o presentarlos con más de 30 días de retraso)', 'Arts. 53 fr. VI y 54 fr. III', '10,000 a 65,000 UMA, o 10% a 100% del valor de la operación, la que sea mayor', rango(10000, 65000)]);
  }
  if (efectivo) sanciones.push([`Aceptar efectivo por encima de los umbrales prohibidos (en inmuebles, desde 8,025 UMA, aprox. ${mx(8025 * uma)})`, 'Arts. 32, 53 fr. VII y 54 fr. III', '10,000 a 65,000 UMA, o 10% a 100% del valor de la operación, la que sea mayor', rango(10000, 65000)]);

  // Barra del índice
  const W = 300, segs = [['#3F8F5A', 0.2], ['#D9B443', 0.2], ['#D9822B', 0.2], ['#B3322B', 0.4]];
  let x0 = 0;
  const barra = segs.map(([c, f]) => { const s = { type: 'rect', x: x0, y: 4, w: W * f, h: 7, color: c }; x0 += W * f; return s; });
  barra.push({ type: 'rect', x: Math.max(0, (W * r.indice) / 100 - 1.5), y: 0, w: 3, h: 15, color: NAVY });

  const meta = (k, v) => ({ text: [{ text: k + ': ', bold: true, color: NAVY }, v || '—'], fontSize: 9.5, margin: [0, 1, 0, 1] });

  const pagina1 = [
    kick('CERTUS PLD · DIAGNÓSTICO PRELIMINAR'),
    h1('Diagnóstico preliminar de exposición PLD/FT'),
    { columns: [
      { stack: [meta('Cliente', cliente.La_Empresa), meta('Contacto', cliente.NombreRepresentanteLegal), meta('Tipo de persona', r.pf ? 'Física' : 'Moral')] },
      { stack: [meta('Fecha', fecha), meta('Folio', folio), meta('Actividades', cliente.Actividades)] }
    ], columnGap: 20 },
    {
      table: { widths: [95, '*'], body: [[
        { text: [{ text: String(r.indice), fontSize: 40, bold: true, font: 'Times', color: NAVY }, { text: ' /100', fontSize: 13, color: MUTED }], margin: [6, 10, 0, 0] },
        { stack: [
          { columns: [{ text: 'Exposición preliminar:', bold: true, fontSize: 12, width: 'auto' }, { ...chip(r.color, r.nivel), width: 'auto', margin: [6, 0, 0, 0] }] },
          { canvas: barra, margin: [0, 8, 0, 2] },
          { columns: [{ text: '0 · menor exposición', fontSize: 7.5, color: MUTED }, { text: '100 · mayor exposición', fontSize: 7.5, color: MUTED, alignment: 'right', width: W - 150 }], width: W },
          { text: [{ text: 'Perfil de riesgo inherente de la actividad: ' }, { text: r.inherente, bold: true }, ' (volumen de operaciones, uso de efectivo y tipo de clientes).'], fontSize: 9, margin: [0, 6, 0, 0] }
        ], margin: [0, 6, 6, 6] }
      ]] },
      layout: { hLineWidth: () => 0, vLineWidth: (i) => (i === 0 ? 3 : 0), vLineColor: () => '#D4944A', fillColor: () => SOFT },
      margin: [0, 12, 0, 8]
    },
    { text: 'El índice mide qué tanto se aparta la operación actual de las obligaciones de la Ley Federal para la Prevención e Identificación de Operaciones con Recursos de Procedencia Ilícita (LFPIORPI), según la información que usted declaró.' + (r.noSe ? ` Hay ${r.noSe} respuesta${r.noSe > 1 ? 's' : ''} «No sé», que se califican con criterio conservador hasta validarlas.` : ''), fontSize: 9.5 },
    h2('Estado por área'),
    tabla(['*', 130, 120], ['Área', 'Estado preliminar', 'Fundamento'], r.bloques.map(x => [
      BLOQUES[x.b][0],
      x.na ? chip('g', 'No aplica') : chip(x.c, x.t),
      x.na ? 'Art. 20: la persona física cumple de forma directa' : BLOQUES[x.b][1]
    ]))
  ];

  const pagina2 = [
    kick('EXPOSICIÓN SANCIONATORIA'),
    h1('Qué podría costar una revisión de la autoridad'),
    { text: `Multas que establece la LFPIORPI para las áreas donde sus respuestas muestran una brecha. Cada incumplimiento puede sancionarse por separado. Las cantidades en pesos usan una UMA de $${uma.toFixed(2)} y son solo de referencia.`, fontSize: 9.5, margin: [0, 0, 0, 8] },
    sanciones.length
      ? tabla(['*', 95, 110, 90], ['Supuesto', 'Fundamento', 'Multa', 'En pesos'], sanciones.map(s => [s[0], s[1].replace(/ (fr\.|y) /g, ' $1\u00A0').replace(/fr\. /g, 'fr.\u00A0'), s[2], { text: s[3], alignment: 'right' }]))
      : caja('Con sus respuestas no se identifican supuestos de sanción en este momento. La sesión de diagnóstico lo confirmará con documentos.'),
    !r.pf && R.rep !== 'si' ? caja({ text: [{ text: 'Sin Representante designado o vigente: ', bold: true }, 'mientras no exista designación aceptada, las obligaciones de la Ley recaen directamente en el órgano de administración o en el administrador único (Art. 20, segundo párrafo).'] }, true) : '',
    fedatario ? caja({ text: [{ text: 'Fedatarios públicos: ', bold: true }, 'la reincidencia o la omisión de Avisos se consideran «notorias deficiencias» y se informan a la autoridad que supervisa la fe pública para que inicie su propio procedimiento (Art. 58).'] }, true) : '',
    String(perfil.historial || '').startsWith('S') ? caja({ text: [{ text: 'Operaciones anteriores: ', bold: true }, 'usted indicó operaciones previas sin identificar clientes o sin presentar Avisos. Conviene evaluarlas antes de cualquier revisión, dentro de la ventana que se explica abajo.'] }, true) : '',
    h2('Su ventana de regularización'),
    perfil.verificacion === 'si'
      ? caja({ text: ['Usted indicó que ya existe una visita o requerimiento. El beneficio del Art. 55 exige regularizarse ', { text: 'antes', bold: true }, ' de que inicien las facultades de verificación, por lo que conviene revisar su situación de inmediato.'] }, true)
      : caja({ text: ['Si se regulariza ', { text: 'de forma espontánea, antes de que la autoridad inicie facultades de verificación', bold: true }, ', y reconoce la falta, la SHCP se abstiene de sancionar por única ocasión. Si ya usó ese beneficio, la multa puede reducirse hasta en 50% (Art. 55).' + (perfil.verificacion === 'ns' ? ' Usted respondió «No sé» sobre requerimientos en curso: lo confirmaremos en la sesión.' : '')] })
  ];

  const pagina3 = [
    kick('SIGUIENTE PASO'),
    h1('Su diagnóstico completo se entrega en una sesión personal'),
    { text: ['Este resultado es una primera lectura. El diagnóstico completo, con la revisión de documentos, el detalle de cada brecha y la ruta de regularización, se entrega ', { text: 'en persona o por videollamada', bold: true }, ' (45 minutos). En esa sesión resolveremos estos puntos:'], fontSize: 10.5 },
    { ol: r.pendientes.slice(), fontSize: 10, margin: [0, 6, 0, 10] },
    {
      table: { widths: ['*'], body: [[{ text: [
        { text: 'Agende su sesión: ', bold: true, color: '#F1C38C' },
        agenda ? { text: agenda, link: agenda, color: '#FFFFFF', decoration: 'underline' } : 'responda este correo y le propondremos horarios',
        ' · contacto@certusconsultores.com.mx · +52 999 163 1363'
      ], color: '#FFFFFF', fontSize: 10, margin: [8, 8, 8, 8] }]] },
      layout: { hLineWidth: () => 0, vLineWidth: () => 0, fillColor: () => NAVY }
    },
    { text: 'Para aprovechar la sesión, tenga a la mano, si existen: acuse de alta en el SPPLD, designación del Representante, manual vigente, acuses de Avisos de los últimos 12 meses y dos expedientes de clientes.', fontSize: 9.5, margin: [0, 10, 0, 0] },
    { canvas: [{ type: 'line', x1: 0, y1: 0, x2: 468, y2: 0, lineWidth: 0.6, lineColor: LINE }], margin: [0, 14, 0, 6] },
    { text: [{ text: 'Alcance de este documento. ', bold: true }, 'Resultado orientativo, generado automáticamente a partir de la información que declaró el cliente, sin revisión documental. No constituye dictamen, opinión jurídica ni garantía de cumplimiento, y no debe usarse ante autoridad alguna. Las referencias corresponden a la LFPIORPI (última reforma DOF 16-07-2025).'], fontSize: 8, color: '#5D6372' },
    { text: [{ text: 'Datos personales. ', bold: true }, constanciaPrivacidad(datos.privacidad) + ' Certus Consultores trata la información recabada conforme a su aviso de privacidad (certusconsultores.com.mx/aviso-de-privacidad), únicamente para los fines ahí señalados. Puede ejercer sus derechos ARCO en contacto@certusconsultores.com.mx.'], fontSize: 8, color: '#5D6372', margin: [0, 4, 0, 0] }
  ];

  const esquinas = (pageSize) => {
    const out = [];
    if (LOGO_E) out.push({ image: LOGO_E, width: 40, absolutePosition: { x: 26, y: 20 } }, { image: LOGO_E, width: 40, absolutePosition: { x: 26, y: pageSize.height - 58 } });
    if (LOGO_C) out.push({ image: LOGO_C, width: 100, absolutePosition: { x: pageSize.width - 126, y: 24 } }, { image: LOGO_C, width: 100, absolutePosition: { x: pageSize.width - 126, y: pageSize.height - 56 } });
    return out;
  };

  const dd = {
    pageSize: 'LETTER',
    pageMargins: [72, 84, 72, 72],
    info: { title: 'Diagnóstico preliminar PLD/FT — ' + (cliente.La_Empresa || ''), author: 'Certus Consultores' },
    defaultStyle: { font: 'Helvetica', fontSize: 10, color: INK, lineHeight: 1.25 },
    background: (page, pageSize) => esquinas(pageSize),
    footer: (page, total) => ({ text: `Folio ${folio} · Página ${page} de ${total} · Documento orientativo`, alignment: 'center', fontSize: 7.5, color: '#7A7F8C', margin: [0, 30, 0, 0] }),
    content: [...pagina1, { text: '', pageBreak: 'after' }, ...pagina2, { text: '', pageBreak: 'after' }, ...pagina3]
  };

  return new Promise((resolve, reject) => {
    try {
      const doc = printer.createPdfKitDocument(dd);
      const chunks = [];
      doc.on('data', c => chunks.push(c));
      doc.on('end', () => resolve({ resultado: { ...r, folio }, buffer: Buffer.concat(chunks) }));
      doc.on('error', reject);
      doc.end();
    } catch (e) { reject(e); }
  });
}

module.exports = { calcular, generarPDF, correoCliente, PREGUNTAS };
