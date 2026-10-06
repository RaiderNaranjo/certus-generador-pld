const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  HeadingLevel, LevelFormat
} = require('docx');
 
// ─────────────────────────────────────────────
// COLORES CERTUS
// ─────────────────────────────────────────────
const NAVY  = '1F3864';
const GOLD  = 'C9A84C';
const WHITE = 'FFFFFF';
const LIGHT = 'F0F4FF';
const GRAY  = '6B7280';
 
// ─────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────
const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const thinBorder = { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' };
const thinBorders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
 
function spacer(pts = 6) {
  return new Paragraph({ spacing: { before: 0, after: pts * 20 }, children: [] });
}
 
function cellText(text, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: 0, after: 0 },
    children: [new TextRun({
      text,
      font: 'Arial',
      size: opts.size || 20,
      bold: opts.bold || false,
      color: opts.color || '1A1A1A',
    })]
  });
}
 
// ─────────────────────────────────────────────
// HEADER NAVY/GOLD
// ─────────────────────────────────────────────
function buildHeader(empresa, rfc, fechaStr) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [6200, 3160],
    rows: [
      new TableRow({
        children: [
          // Celda izquierda — CERTUS PLD + título
          new TableCell({
            borders: noBorders,
            shading: { fill: NAVY, type: ShadingType.CLEAR },
            margins: { top: 200, bottom: 200, left: 240, right: 120 },
            verticalAlign: VerticalAlign.CENTER,
            width: { size: 6200, type: WidthType.DXA },
            children: [
              new Paragraph({
                spacing: { before: 0, after: 40 },
                children: [new TextRun({ text: 'CERTUS PLD', font: 'Arial', size: 28, bold: true, color: GOLD })]
              }),
              new Paragraph({
                spacing: { before: 0, after: 20 },
                children: [new TextRun({ text: 'Cronograma de Implementación', font: 'Arial', size: 22, bold: true, color: WHITE })]
              }),
              new Paragraph({
                spacing: { before: 0, after: 0 },
                children: [new TextRun({ text: 'Programa de Cumplimiento PLD/FT — 90 Días', font: 'Arial', size: 18, color: 'AAAAAA' })]
              }),
            ]
          }),
          // Celda derecha — datos del cliente
          new TableCell({
            borders: noBorders,
            shading: { fill: '162D52', type: ShadingType.CLEAR },
            margins: { top: 200, bottom: 200, left: 200, right: 200 },
            verticalAlign: VerticalAlign.CENTER,
            width: { size: 3160, type: WidthType.DXA },
            children: [
              new Paragraph({
                spacing: { before: 0, after: 60 },
                children: [new TextRun({ text: empresa, font: 'Arial', size: 20, bold: true, color: WHITE })]
              }),
              new Paragraph({
                spacing: { before: 0, after: 40 },
                children: [new TextRun({ text: `RFC: ${rfc}`, font: 'Arial', size: 17, color: GOLD })]
              }),
              new Paragraph({
                spacing: { before: 0, after: 0 },
                children: [new TextRun({ text: `Inicio: ${fechaStr}`, font: 'Arial', size: 16, color: 'AAAAAA' })]
              }),
            ]
          }),
        ]
      })
    ]
  });
}
 
// ─────────────────────────────────────────────
// FILA DE FASE (encabezado de semana)
// ─────────────────────────────────────────────
function buildFaseRow(titulo, fecha) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [6840, 2520],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: noBorders,
            shading: { fill: NAVY, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 200, right: 120 },
            width: { size: 6840, type: WidthType.DXA },
            children: [cellText(titulo, { bold: true, color: WHITE, size: 20 })]
          }),
          new TableCell({
            borders: noBorders,
            shading: { fill: GOLD, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 160, right: 160 },
            verticalAlign: VerticalAlign.CENTER,
            width: { size: 2520, type: WidthType.DXA },
            children: [cellText(fecha, { bold: true, color: NAVY, size: 18, align: AlignmentType.CENTER })]
          }),
        ]
      })
    ]
  });
}
 
// ─────────────────────────────────────────────
// TABLA DE DOCUMENTOS DE UNA FASE
// ─────────────────────────────────────────────
function buildDocTable(docs, objetivo) {
  const headerRow = new TableRow({
    children: [
      new TableCell({
        borders: noBorders,
        shading: { fill: LIGHT, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 160, right: 80 },
        width: { size: 1440, type: WidthType.DXA },
        children: [cellText('Documento', { bold: true, color: NAVY, size: 18 })]
      }),
      new TableCell({
        borders: noBorders,
        shading: { fill: LIGHT, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 80 },
        width: { size: 2800, type: WidthType.DXA },
        children: [cellText('Nombre', { bold: true, color: NAVY, size: 18 })]
      }),
      new TableCell({
        borders: noBorders,
        shading: { fill: LIGHT, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        width: { size: 5120, type: WidthType.DXA },
        children: [cellText('Descripción', { bold: true, color: NAVY, size: 18 })]
      }),
    ]
  });
 
  const docRows = docs.map((doc, i) =>
    new TableRow({
      children: [
        new TableCell({
          borders: noBorders,
          shading: { fill: i % 2 === 0 ? 'FFFFFF' : 'F8F7F5', type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 160, right: 80 },
          width: { size: 1440, type: WidthType.DXA },
          children: [cellText(doc.id, { bold: true, color: NAVY, size: 19 })]
        }),
        new TableCell({
          borders: noBorders,
          shading: { fill: i % 2 === 0 ? 'FFFFFF' : 'F8F7F5', type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 80 },
          width: { size: 2800, type: WidthType.DXA },
          children: [cellText(doc.nombre, { bold: false, color: '1A1A1A', size: 19 })]
        }),
        new TableCell({
          borders: noBorders,
          shading: { fill: i % 2 === 0 ? 'FFFFFF' : 'F8F7F5', type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          width: { size: 5120, type: WidthType.DXA },
          children: [cellText(doc.descripcion, { color: GRAY, size: 18 })]
        }),
      ]
    })
  );
 
  // Fila de objetivo
  const objRow = new TableRow({
    children: [
      new TableCell({
        borders: noBorders,
        columnSpan: 3,
        shading: { fill: 'FEF9EC', type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 160, right: 120 },
        width: { size: 9360, type: WidthType.DXA },
        children: [
          new Paragraph({
            spacing: { before: 0, after: 0 },
            children: [
              new TextRun({ text: 'OBJETIVO: ', font: 'Arial', size: 18, bold: true, color: NAVY }),
              new TextRun({ text: objetivo, font: 'Arial', size: 18, color: '92400E' }),
            ]
          })
        ]
      })
    ]
  });
 
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [1440, 2800, 5120],
    rows: [headerRow, ...docRows, objRow]
  });
}
 
// ─────────────────────────────────────────────
// TABLA DE OBSERVACIONES / PRÓXIMOS PASOS
// ─────────────────────────────────────────────
function buildInfoTable(items, titulo, colorFondo) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: noBorders,
            shading: { fill: colorFondo, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 60, left: 200, right: 200 },
            width: { size: 9360, type: WidthType.DXA },
            children: [cellText(titulo, { bold: true, color: NAVY, size: 20 })]
          })
        ]
      }),
      ...items.map((item, i) => new TableRow({
        children: [
          new TableCell({
            borders: noBorders,
            shading: { fill: i % 2 === 0 ? 'FFFFFF' : 'F8F7F5', type: ShadingType.CLEAR },
            margins: { top: 70, bottom: 70, left: 200, right: 200 },
            width: { size: 9360, type: WidthType.DXA },
            children: [cellText(`${i + 1}.  ${item}`, { color: '1A1A1A', size: 19 })]
          })
        ]
      }))
    ]
  });
}
 
// ─────────────────────────────────────────────
// FOOTER CERTUS
// ─────────────────────────────────────────────
function buildFooter() {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: noBorders,
            shading: { fill: NAVY, type: ShadingType.CLEAR },
            margins: { top: 160, bottom: 160, left: 240, right: 240 },
            width: { size: 9360, type: WidthType.DXA },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 40 },
                children: [new TextRun({ text: 'Certus Consultores', font: 'Arial', size: 22, bold: true, color: GOLD })]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 30 },
                children: [new TextRun({ text: 'Especialistas en Cumplimiento Normativo PLD/FT', font: 'Arial', size: 18, color: 'AAAAAA' })]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 0 },
                children: [new TextRun({ text: 'Mérida, Yucatán  •  contacto@certusconsultores.com.mx  •  Tel: 999 163 1363', font: 'Arial', size: 17, color: 'AAAAAA' })]
              }),
            ]
          })
        ]
      })
    ]
  });
}
 
// ─────────────────────────────────────────────
// FUNCIÓN PRINCIPAL — GENERA EL .docx
// ─────────────────────────────────────────────
async function generarCronograma(empresa, rfc, opts = {}) {
  const hoy = new Date();
  const fmt = (d) => d.toLocaleDateString('es-MX', { day: '2-digit', month: 'short', year: 'numeric' });
  const fmtLargo = (d) => d.toLocaleDateString('es-MX', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
 
  const s = (weeks) => { const d = new Date(hoy); d.setDate(d.getDate() + weeks * 7); return d; };
 
  const fases = [
    {
      titulo: 'SEMANA 1 — Documentos Fundacionales',
      fecha: fmt(s(1)),
      objetivo: 'Establecer la base legal y designar al Representante de Cumplimiento',
      docs: [
        { id: 'D0', nombre: 'Carta de Recomendaciones Urgentes', descripcion: 'Comunicado inicial con recomendaciones críticas de cumplimiento' },
        { id: 'D1', nombre: 'Dictamen de Sujeción Normativa', descripcion: 'Evaluación legal de la obligación LFPIORPI' },
        opts.tipoPersona === 'PF'
          ? { id: 'D2', nombre: 'Constancia de Asunción Personal', descripcion: 'El profesional asume personal y directamente el cumplimiento (art. 20 LFPIORPI)' }
          : { id: 'D2', nombre: 'Acta de Designación del RC', descripcion: 'Documento de formalización del Representante de Cumplimiento' },
        { id: 'D3', nombre: 'Expediente de Registro SPPLD', descripcion: 'Documentos para registrar ante el SAT' },
      ]
    },
    {
      titulo: 'SEMANA 2 — Análisis y Procedimientos',
      fecha: fmt(s(2)),
      objetivo: 'Evaluar brechas normativas y establecer metodología de riesgo',
      docs: [
        { id: 'D4', nombre: 'Gap Analysis PLD', descripcion: 'Análisis de brechas normativas de la empresa' },
        { id: 'D5', nombre: 'Manual de Cumplimiento PLD/FT', descripcion: 'Procedimientos y políticas de cumplimiento' },
        { id: 'D6', nombre: 'Metodología EBR', descripcion: 'Sistema de Enfoque Basado en Riesgo del cliente' },
      ]
    },
    {
      titulo: 'SEMANA 3 — Conocimiento del Cliente (KYC)',
      fecha: fmt(s(3)),
      objetivo: 'Implementar procesos de identificación, monitoreo y reporte',
      docs: [
        { id: 'D7', nombre: 'Expediente KYC', descripcion: 'Procesos de identificación de clientes' },
        { id: 'D8', nombre: 'Política de PEPs', descripcion: 'Procedimientos para detectar Personas Expuestas Políticamente' },
        { id: 'D9', nombre: 'Reporte de Operación Inusual (ROI)', descripcion: 'Protocolo de alertas y reportes al SAT' },
        { id: 'D10', nombre: 'Política de Resguardo Documental', descripcion: 'Estándares de conservación de documentos' },
      ]
    },
    {
      titulo: 'SEMANA 4 — Capacitación y Control',
      fecha: fmt(s(4)),
      objetivo: 'Capacitar al personal y establecer controles operativos',
      docs: [
        { id: 'D11',  nombre: 'Programa Anual de Capacitación', descripcion: 'Plan de capacitación del personal' },
        { id: 'D11B', nombre: 'Casos Prácticos Certus PLD', descripcion: 'Ejemplos de aplicación práctica' },
        { id: 'D11C', nombre: 'Evaluación de Comprensión', descripcion: 'Prueba de conocimiento del personal' },
        { id: 'D12',  nombre: 'Control de Avisos SAT', descripcion: 'Sistema de control de reportes mensuales' },
      ]
    },
    {
      titulo: 'SEMANA 8 — Auditoría y Cierre',
      fecha: fmt(s(8)),
      objetivo: 'Validar la implementación completa y generar informe final',
      docs: [
        { id: 'D13', nombre: 'Informe de Auditoría PLD', descripcion: 'Validación de implementación y hallazgos' },
      ]
    },
  ];
 
  const observaciones = [
    'Todos los documentos son personalizados para la empresa cliente',
    'El Representante de Cumplimiento debe estar disponible para las capacitaciones',
    'Se requieren 2–4 horas semanales de dedicación interna del cliente',
    'Las reuniones de seguimiento se recomiendan cada dos semanas',
    'El cliente debe completar el registro en SPPLD antes de la Semana 2',
  ];
 
  const proximosPasos = [
    'Revisar los documentos entregados en Semana 1',
    'Designar personal clave para cada área de cumplimiento',
    'Agendar sesión de capacitación para el personal (Semana 2)',
    'Confirmar que el SPPLD está correctamente registrado (antes de Semana 2)',
    'Establecer el sistema de reportes mensuales SAT (Semana 3)',
  ];
 
  // ── CONSTRUIR DOCUMENTO ──
  const children = [];
 
  // Header
  children.push(buildHeader(empresa, rfc, fmtLargo(hoy)));
  children.push(spacer(14));
 
  // Fases
  for (const fase of fases) {
    children.push(buildFaseRow(fase.titulo, fase.fecha));
    children.push(buildDocTable(fase.docs, fase.objetivo));
    children.push(spacer(10));
  }
 
  // Observaciones y próximos pasos
  children.push(buildInfoTable(observaciones, 'OBSERVACIONES IMPORTANTES', 'EFF6FF'));
  children.push(spacer(8));
  children.push(buildInfoTable(proximosPasos, 'PRÓXIMOS PASOS', 'F0FDF4'));
  children.push(spacer(14));
 
  // Footer
  children.push(buildFooter());
 
  const doc = new Document({
    styles: {
      default: { document: { run: { font: 'Arial', size: 20 } } }
    },
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 720, right: 720, bottom: 720, left: 720 }
        }
      },
      children
    }]
  });
 
  return await Packer.toBuffer(doc);
}
 
// ─────────────────────────────────────────────
// MAIN
// ─────────────────────────────────────────────
async function main() {
  try {
    let datosFormulario = {};
 
    if (process.argv[2]) {
      try {
        datosFormulario = JSON.parse(process.argv[2]);
      } catch (e) {
        console.error('❌ Error parseando JSON.');
        process.exit(1);
      }
    } else if (fs.existsSync('cliente.json')) {
      try {
        const contenido = fs.readFileSync('cliente.json', 'utf8');
        datosFormulario = JSON.parse(contenido);
        console.log('📂 Datos leídos de cliente.json\n');
      } catch (e) {
        console.error('❌ Error leyendo cliente.json');
        process.exit(1);
      }
    } else {
      datosFormulario = {
        'La Empresa': 'Test Inmobiliaria',
        'RFC': 'TST123456789',
        'Ciudad': 'Mérida'
      };
      console.log('⚠️  Sin datos. Usando valores de prueba...\n');
    }
 
    const empresa = datosFormulario['La Empresa'] || 'Cliente';
    const rfc     = datosFormulario['RFC']        || 'RFC';
    const ciudad  = datosFormulario['Ciudad']     || 'Mérida';
 
    console.log('\n🚀 GENERADOR CERTUS PLD\n');
    console.log(`📋 Cliente: ${empresa}`);
    console.log(`🆔 RFC: ${rfc}`);
    console.log(`📍 Ciudad: ${ciudad}\n`);
 
    const alias = { profesionista: 'profesionistas', profesional: 'profesionistas', profesionales: 'profesionistas', notaria: 'notarial', notaría: 'notarial' };
    let tipo = (datosFormulario['Tipo'] || 'inmobiliaria').toString().trim().toLowerCase();
    tipo = alias[tipo] || tipo;
    if (!['inmobiliaria', 'notarial', 'profesionistas'].includes(tipo)) {
      console.error(`❌ Tipo inválido: ${tipo} (use inmobiliaria, notarial o profesionistas)`);
      process.exit(1);
    }
    // Persona física (profesional independiente) o moral (despacho) — solo aplica a profesionistas
    const tpRaw = (datosFormulario['TipoPersona'] || '').toString().trim().toLowerCase();
    const tipoPersona = /^(pf|fisica|física|persona f)/.test(tpRaw) ? 'PF' : (/^(pm|moral|persona m|despacho)/.test(tpRaw) ? 'PM' : '');
    const dirPlantillas = path.join(__dirname, 'plantillas', tipo);
    if (!fs.existsSync(dirPlantillas)) {
      console.error(`❌ Carpeta plantillas no encontrada en: ${dirPlantillas}`);
      process.exit(1);
    }

    const archivos = fs.readdirSync(dirPlantillas)
      .filter(f => f.endsWith('.docx') && !f.startsWith('~$'))
      .sort();
    // Si hay varias versiones del mismo documento (_v2, _v3), solo la más reciente
    const vigentes = {};
    for (const f of archivos) {
      const m = /^(.*?)_v(\d+)\.docx$/.exec(f);
      const base = m ? m[1].replace(/_CertusPLD$/, '') : f;
      const ver = m ? Number(m[2]) : 0;
      if (!vigentes[base] || ver > vigentes[base].ver) vigentes[base] = { f, ver };
    }
    const seleccion = new Set(Object.values(vigentes).map(v => v.f));
    archivos.splice(0, archivos.length, ...archivos.filter(f => seleccion.has(f)));

    // Profesionistas: solo el documento D2 que corresponde al tipo de sujeto obligado
    if (tipo === 'profesionistas' && tipoPersona) {
      const excluir = tipoPersona === 'PF' ? 'D2-PROF-PM' : 'D2-PROF-PF';
      archivos.splice(0, archivos.length, ...archivos.filter(f => !f.startsWith(excluir)));
    }

    console.log(`📄 Personalizando ${archivos.length} documentos (${tipo}${tipoPersona ? ', ' + tipoPersona : ''})...\n`);

    const docsPersonalizados = {};

    const domicilio = datosFormulario['Domicilio'] || (ciudad + ', México');
    const rc = datosFormulario['RepresentantePLD'] || '';
    const repLegal = datosFormulario['NombreRepresentanteLegal'] || '';
    const actividades = datosFormulario['Actividades'] ||
      (tipo === 'notarial' ? 'Fe pública notarial'
        : tipo === 'profesionistas' ? 'Servicios profesionales independientes (art. 17, fr. XI, LFPIORPI)'
        : 'Compraventa de inmuebles, arrendamiento y desarrollo de proyectos inmobiliarios');
    const hoy = new Date();
    const fechaCorta = hoy.toLocaleDateString('es-MX');
    const fechaLarga = hoy.toLocaleDateString('es-MX', { year: 'numeric', month: 'long', day: 'numeric' });

    // Solo se reemplaza lo que se conoce; el resto queda como campo por llenar
    const reemplazos = {
      '[NOMBRE DE LA INMOBILIARIA]': empresa,
      '[Nombre de la Inmobiliaria]': empresa,
      '[NOMBRE DEL CLIENTE]': empresa,
      '[La_Empresa]': empresa,
      '[La Empresa]': empresa,
      '[EMPRESA]': empresa,
      '[Empresa]': empresa,
      '[Razón social completa]': empresa,
      '[Denominación de la Notaría]': empresa,
      '[Denominación completa]': empresa,
      '[NOMBRE DE LA NOTARIA]': empresa,
      '[RFC]': rfc,
      '[RFC_EMPRESA]': rfc,
      '[RFC de la empresa]': rfc,
      '[CIUDAD]': ciudad,
      '[Ciudad]': ciudad,
      '[Domicilio]': domicilio,
      '[DOMICILIO]': domicilio,
      '[DOMICILIO FISCAL]': domicilio,
      '[Calle, Numero, Colonia, CP, Municipio, Estado]': domicilio,
      '[FECHA]': fechaCorta,
      '[Fecha]': fechaCorta,
      '[FECHACOMPLETA]': fechaLarga,
      '[ACTIVIDADES]': actividades,
      '[VERSION]': '1.0'
    };
    if (rc) {
      Object.assign(reemplazos, {
        '[RC_Nombre]': rc, '[NOMBRE RC]': rc,
        '[Nombre del Representante de Cumplimiento]': rc, '[Nombre del RC]': rc
      });
    }
    if (repLegal) {
      Object.assign(reemplazos, {
        '[NOMBRE DEL REPRESENTANTE LEGAL]': repLegal,
        '[Nombre del Notario Titular]': repLegal,
        '[NOTARIO TITULAR]': repLegal
      });
    }

    if (tipo === 'profesionistas') {
      Object.assign(reemplazos, {
        '[Nombre]': empresa,
        '[Nombre del despacho]': empresa,
        '[Razón social completa del despacho]': empresa,
        '[Nombre completo del profesional o denominación del despacho]': empresa,
        '[Nombre del profesional / representante del despacho]': repLegal || empresa,
        '[Nombre completo]': empresa,
        '[Nombre del profesional]': empresa,
        '[Ciudad, Fecha]': `${ciudad}, ${fechaLarga}`
      });
      if (tipoPersona) {
        reemplazos['[Persona física — profesional independiente  /  Persona moral — despacho]'] =
          tipoPersona === 'PF' ? 'Persona física — profesional independiente' : 'Persona moral — despacho';
        reemplazos['[profesional / despacho]'] = tipoPersona === 'PF' ? 'profesional' : 'despacho';
        const titular = tipoPersona === 'PF' ? 'Profesional Titular' : null;
        reemplazos['[Profesional Titular / Órgano de administración]'] = titular || 'Órgano de administración';
        reemplazos['[Profesional Titular / Órgano de administración del despacho]'] = titular || 'Órgano de administración del despacho';
        reemplazos['[Profesional Titular / Repr. de Cumplimiento]'] = titular || 'Representante de Cumplimiento';
        reemplazos['[Profesional Titular / Representante de Cumplimiento]'] = titular || 'Representante de Cumplimiento';
      }
      if (datosFormulario['Actividades']) {
        reemplazos['[Abogacía / Contaduría pública / Asesoría corporativa — especificar]'] = actividades;
        reemplazos['[Actividad profesional]'] = actividades;
      }
    }

    // Incisos que no se prestan / no se formalizan (según intake)
    const incNo = (datosFormulario['IncisosNoPrestados'] || '').toString().trim();
    if (incNo) {
      if (tipo === 'notarial') {
        reemplazos['[INCISOS QUE LA NOTARÍA NO FORMALIZA, O “NINGUNO”]'] = incNo;
        reemplazos['[INCISOS QUE LA NOTARÍA NO FORMALIZA, O "NINGUNO"]'] = incNo;
      } else if (tipo === 'profesionistas') {
        reemplazos['[INCISOS QUE NO SE PRESTAN — a), b), c), d) y/o e)]'] = incNo;
      }
    }
    if (tipo === 'profesionistas') {
      const incl = (datosFormulario['IncisosTexto'] || '').toString().trim();
      const excl = (datosFormulario['LineasExcluidas'] || '').toString().trim();
      if (datosFormulario['Incisos'] !== undefined && datosFormulario['IncisosNoPrestados']) {
        reemplazos['[LISTA DE LÍNEAS INCLUIDAS]'] = incl || 'ninguna';
        reemplazos['[LISTA DE LÍNEAS EXCLUIDAS]'] = excl || 'las demás líneas de servicio que no involucran los incisos a) a e)';
      }
    }

    // Reemplazo a nivel de párrafo: cubre marcadores partidos en varios runs
    const reemplazarEnParrafos = (xml) => xml.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, (p) => {
      const ts = [...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)];
      if (ts.length < 2) return p;
      const texto = ts.map(m => m[1]).join('');
      if (!texto.includes('[')) return p;
      let nuevo = texto;
      for (const [ph, val] of Object.entries(reemplazos)) {
        const phx = ph.replace(/&/g, '&amp;');
        nuevo = nuevo.split(phx).join(String(val).replace(/&/g, '&amp;').replace(/</g, '&lt;'));
      }
      if (nuevo === texto) return p;
      let i = 0;
      return p.replace(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g, () =>
        i++ === 0 ? `<w:t xml:space="preserve">${nuevo}</w:t>` : '<w:t></w:t>');
    });

    // Cronograma propio del producto (p. ej. CRONOGRAMA-NOT): fechas por semana, titular e incisos
    const esCronogramaPropio = (a) => /^CRONOGRAMA-/i.test(a);
    const tieneCronogramaPropio = archivos.some(esCronogramaPropio);
    const fmtFecha = (d) => d.toLocaleDateString('es-MX', { day: '2-digit', month: 'short', year: 'numeric' });
    const sumaSemanas = (w) => { const d = new Date(hoy); d.setDate(d.getDate() + w * 7); return d; };
    const personalizarCronograma = (xml) => {
      const fechas = [hoy, sumaSemanas(1), sumaSemanas(2), sumaSemanas(3), sumaSemanas(4), sumaSemanas(8)].map(fmtFecha);
      const letras = (datosFormulario['Incisos'] || '').toString().split(',').map(x => x.trim().toLowerCase()).filter(Boolean);
      const titular = repLegal || empresa;
      let idx = 0;
      return xml.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, (p) => {
        const ts = [...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)];
        if (!ts.length) return p;
        const texto = ts.map(m => m[1]).join('');
        let nuevo = texto;
        if (nuevo.includes('[Fecha]')) {
          nuevo = nuevo.split('[Fecha]').join(fechas[Math.min(idx++, fechas.length - 1)]);
        }
        if (nuevo.includes('[Nombre]')) nuevo = nuevo.split('[Nombre]').join(titular.replace(/&/g, '&amp;'));
        if (/\[ \] a/.test(nuevo) && letras.length) {
          nuevo = nuevo.replace(/\[ \] ([a-e])/g, (m, l) => (letras.includes(l) ? '[X] ' : '[ ] ') + l);
        }
        if (nuevo === texto) return p;
        let i = 0;
        return p.replace(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g, () =>
          i++ === 0 ? `<w:t xml:space="preserve">${nuevo}</w:t>` : '<w:t></w:t>');
      });
    };

    for (const archivo of archivos) {
      const rutaPlantilla = path.join(dirPlantillas, archivo);
      const buffer = fs.readFileSync(rutaPlantilla);

      const zip = new JSZip();
      await zip.loadAsync(buffer);

      for (const parte of Object.keys(zip.files).filter(n => /^word\/(document|header\d*|footer\d*)\.xml$/.test(n))) {
        let xml = await zip.file(parte).async('string');
        if (esCronogramaPropio(archivo) && parte === 'word/document.xml') xml = personalizarCronograma(xml);
        for (const [placeholder, valor] of Object.entries(reemplazos)) {
          xml = xml.split(placeholder).join(String(valor).replace(/&/g, '&amp;').replace(/</g, '&lt;'));
        }
        xml = reemplazarEnParrafos(xml);
        zip.file(parte, xml);
      }
      const personalizado = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });

      docsPersonalizados[archivo] = personalizado;
      console.log(`  ✓ ${archivo}`);
    }

    let cronogramaBuffer = null;
    let nombreCronograma = null;
    if (!tieneCronogramaPropio) {
      console.log(`\n📅 Generando cronograma...\n`);
      cronogramaBuffer = await generarCronograma(empresa, rfc, { tipo, tipoPersona });
      nombreCronograma = `CRONOGRAMA_${empresa.replace(/\s/g, '_')}_${Date.now()}.docx`;
      fs.writeFileSync(nombreCronograma, cronogramaBuffer);
      console.log(`✓ Cronograma creado: ${nombreCronograma}\n`);
    }
 
    console.log(`\n📦 Comprimiendo documentos...\n`);
    const zipFinal = new JSZip();
    for (const [nombre, buffer] of Object.entries(docsPersonalizados)) {
      zipFinal.file(nombre, buffer);
    }
    if (cronogramaBuffer) zipFinal.file(nombreCronograma, cronogramaBuffer);
 
    const zipBuffer = await zipFinal.generateAsync({ type: 'nodebuffer' });
    const nombreZip = `CERTUS_PLD_${empresa.replace(/\s/g, '_')}_${Date.now()}.zip`;
    fs.writeFileSync(nombreZip, zipBuffer);
    console.log(`✓ ZIP creado: ${nombreZip}`);
    console.log(`  Tamaño: ${(zipBuffer.length / 1024 / 1024).toFixed(2)} MB\n`);
 
    console.log(`✅ ¡COMPLETADO!\n`);
    console.log(`📍 Archivos generados:`);
    console.log(`   • ${nombreZip}`);
    if (nombreCronograma) console.log(`   • ${nombreCronograma}`);
    console.log('');
    console.log(`💡 Próxima vez: node generador.js '{"La Empresa":"Nuevo Cliente","RFC":"XYZ123","Ciudad":"Mérida"}'\n`);
 
  } catch (error) {
    console.error('\n❌ ERROR:', error.message);
    console.error(error);
    process.exit(1);
  }
}
 
main();
