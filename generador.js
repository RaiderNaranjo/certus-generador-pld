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
async function generarCronograma(empresa, rfc) {
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
        { id: 'D2', nombre: 'Acta de Designación del RC', descripcion: 'Documento de formalización del Representante de Cumplimiento' },
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
 
    const dirPlantillas = path.join(__dirname, 'plantillas');
    if (!fs.existsSync(dirPlantillas)) {
      console.error(`❌ Carpeta plantillas no encontrada en: ${dirPlantillas}`);
      process.exit(1);
    }
 
    const archivos = fs.readdirSync(dirPlantillas)
      .filter(f => f.endsWith('.docx'))
      .sort();
 
    console.log(`📄 Personalizando ${archivos.length} documentos...\n`);
 
    const docsPersonalizados = {};
 
    const reemplazos = {
      '[NOMBRE DE LA INMOBILIARIA]': empresa,
      '[Nombre de la Inmobiliaria]': empresa,
      '[NOMBRE DEL CLIENTE]': empresa,
      '[RFC]': rfc,
      '[RFC de la empresa]': rfc,
      '[CIUDAD]': ciudad,
      '[DOMICILIO]': ciudad + ', México',
      '[DOMICILIO FISCAL]': ciudad + ', México',
      '[FECHA]': new Date().toLocaleDateString('es-MX'),
      '[FECHACOMPLETA]': new Date().toLocaleDateString('es-MX', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' }),
      '[ACTIVIDADES]': 'Compraventa de inmuebles, arrendamiento y desarrollo de proyectos inmobiliarios',
      '[NOMBRE DEL REPRESENTANTE LEGAL]': '[Nombre del Representante Legal]',
      '[NOMBRE RC]': '[Nombre del RC]',
      '[CARGO RC]': '[Cargo del RC]',
      '[EMAIL RC]': '[Email del RC]',
      '[TELEFONO RC]': '[Teléfono del RC]',
      '[ÓRGANO DE ADMINISTRACIÓN]': '[Órgano de administración]',
      '[NUMERO OPERACIONES]': '[Número de operaciones]',
      '[VERSION]': '1.0',
      '[Razón social completa]': empresa,
      '[Calle, Numero, Colonia, CP, Municipio, Estado]': ciudad + ', México',
      '[Número asignado por el SAT]': '[SPPLD]',
      '[SPPLD]': '[SPPLD]'
    };
 
    for (const archivo of archivos) {
      const rutaPlantilla = path.join(dirPlantillas, archivo);
      const buffer = fs.readFileSync(rutaPlantilla);
 
      const zip = new JSZip();
      await zip.loadAsync(buffer);
 
      let docXml = await zip.file('word/document.xml').async('string');
 
      for (const [placeholder, valor] of Object.entries(reemplazos)) {
        docXml = docXml.split(placeholder).join(valor);
      }
 
      zip.file('word/document.xml', docXml);
      const personalizado = await zip.generateAsync({ type: 'nodebuffer' });
 
      docsPersonalizados[archivo] = personalizado;
      console.log(`  ✓ ${archivo}`);
    }
 
    console.log(`\n📅 Generando cronograma...\n`);
    const cronogramaBuffer = await generarCronograma(empresa, rfc);
    const nombreCronograma = `CRONOGRAMA_${empresa.replace(/\s/g, '_')}_${Date.now()}.docx`;
    fs.writeFileSync(nombreCronograma, cronogramaBuffer);
    console.log(`✓ Cronograma creado: ${nombreCronograma}\n`);
 
    console.log(`\n📦 Comprimiendo documentos...\n`);
    const zipFinal = new JSZip();
    for (const [nombre, buffer] of Object.entries(docsPersonalizados)) {
      zipFinal.file(nombre, buffer);
    }
    zipFinal.file(nombreCronograma, cronogramaBuffer);
 
    const zipBuffer = await zipFinal.generateAsync({ type: 'nodebuffer' });
    const nombreZip = `CERTUS_PLD_${empresa.replace(/\s/g, '_')}_${Date.now()}.zip`;
    fs.writeFileSync(nombreZip, zipBuffer);
    console.log(`✓ ZIP creado: ${nombreZip}`);
    console.log(`  Tamaño: ${(zipBuffer.length / 1024 / 1024).toFixed(2)} MB\n`);
 
    console.log(`✅ ¡COMPLETADO!\n`);
    console.log(`📍 Archivos generados:`);
    console.log(`   • ${nombreZip}`);
    console.log(`   • ${nombreCronograma}\n`);
    console.log(`💡 Próxima vez: node generador.js '{"La Empresa":"Nuevo Cliente","RFC":"XYZ123","Ciudad":"Mérida"}'\n`);
 
  } catch (error) {
    console.error('\n❌ ERROR:', error.message);
    console.error(error);
    process.exit(1);
  }
}
 
main();
