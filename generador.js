const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

async function generarCronograma(empresa, rfc) {
  const hoy = new Date();
  const semana1 = new Date(hoy);
  semana1.setDate(semana1.getDate() + 7);
  const semana2 = new Date(semana1);
  semana2.setDate(semana2.getDate() + 7);
  const semana3 = new Date(semana2);
  semana3.setDate(semana3.getDate() + 7);
  const semana4 = new Date(semana3);
  semana4.setDate(semana4.getDate() + 7);
  const semana8 = new Date(semana4);
  semana8.setDate(semana8.getDate() + 28);
  
  const cronograma = `CRONOGRAMA DE IMPLEMENTACIÓN
Programa de Cumplimiento PLD/FT — 90 Días

═════════════════════════════════════════════════════════════════

CLIENTE: ${empresa}
RFC: ${rfc}
RESPONSABLE: Certus Consultores
FECHA DE INICIO: ${hoy.toLocaleDateString('es-MX', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}

═════════════════════════════════════════════════════════════════

SEMANA 1 — DOCUMENTOS FUNDACIONALES
Fecha de entrega: ${semana1.toLocaleDateString('es-MX')}

├─ D0: Carta de Recomendaciones Urgentes
│   Descripción: Comunicado inicial con recomendaciones críticas de cumplimiento
│
├─ D1: Dictamen de Sujeción Normativa
│   Descripción: Evaluación legal de la obligación LFPIORPI
│
├─ D2: Acta de Designación del Representante de Cumplimiento
│   Descripción: Documento de formalización del RC
│
└─ D3: Expediente de Registro SPPLD
    Descripción: Documentos para registrar ante SAT

OBJETIVO: Establecer la base legal y designar al RC

───────────────────────────────────────────────────────────────

SEMANA 2 — ANÁLISIS Y PROCEDIMIENTOS
Fecha de entrega: ${semana2.toLocaleDateString('es-MX')}

├─ D4: Gap Analysis PLD
│   Descripción: Análisis de brechas normativas
│
├─ D5: Manual de Cumplimiento PLD/FT
│   Descripción: Procedimientos y políticas de cumplimiento
│
└─ D6: Metodología de Enfoque Basado en Riesgo (EBR)
    Descripción: Sistema de evaluación de riesgo del cliente

OBJETIVO: Evaluar brechas y establecer metodología de riesgo

───────────────────────────────────────────────────────────────

SEMANA 3 — PROCESOS DE CONOCIMIENTO DEL CLIENTE
Fecha de entrega: ${semana3.toLocaleDateString('es-MX')}

├─ D7: Expediente KYC (Know Your Customer)
│   Descripción: Procesos de identificación de clientes
│
├─ D8: Política de PEPs (Personas Expuestas Políticamente)
│   Descripción: Procedimientos para detectar PEPs
│
├─ D9: Reporte de Operación Inusual (ROI)
│   Descripción: Protocolo de alertas y reportes
│
└─ D10: Política de Resguardo Documental
    Descripción: Estándares de conservación de documentos

OBJETIVO: Implementar procesos de identificación y monitoreo

───────────────────────────────────────────────────────────────

SEMANA 4 — CAPACITACIÓN Y CONTROL
Fecha de entrega: ${semana4.toLocaleDateString('es-MX')}

├─ D11: Programa Anual de Capacitación
│   Descripción: Plan de capacitación del personal
│
├─ D11B: Casos Prácticos Certus PLD
│   Descripción: Ejemplos de aplicación práctica
│
├─ D11C: Evaluación de Comprensión
│   Descripción: Prueba de conocimiento del personal
│
└─ D12: Control de Avisos SAT
    Descripción: Sistema de control de reportes mensuales

OBJETIVO: Capacitar personal y establecer controles

───────────────────────────────────────────────────────────────

SEMANA 8 — AUDITORÍA Y CIERRE
Fecha de entrega: ${semana8.toLocaleDateString('es-MX')}

└─ D13: Informe de Auditoría PLD
    Descripción: Validación de implementación y hallazgos

OBJETIVO: Validar implementación y generar informe final

═════════════════════════════════════════════════════════════════

OBSERVACIONES IMPORTANTES:

• Todos los documentos son personalizados para el cliente
• El Representante de Cumplimiento debe estar disponible para capacitaciones
• Se requieren 2-4 horas semanales de dedicación interna
• Las reuniones de seguimiento se recomiendan cada dos semanas
• El cliente debe completar registros en SPPLD antes de Semana 2

PRÓXIMOS PASOS:

1. Revisar los documentos entregados en Semana 1
2. Designar personal clave para cada área de cumplimiento
3. Agendar sesión de capacitación para el personal (Semana 2)
4. Confirmar que el SPPLD está correctamente registrado (antes de Semana 2)
5. Establecer sistema de reportes mensuales SAT (Semana 3)

SOPORTE Y CONTACTO:

Durante la implementación, el cliente tiene acceso a:
• Consultas técnicas vía email
• Sesiones de clarificación por videollamada (máximo 10/año en suscripción)
• Materiales de referencia y guías paso a paso
• Actualizaciones de documentos conforme a cambios normativos

═════════════════════════════════════════════════════════════════

Certus Consultores
Especialistas en Cumplimiento Normativo PLD/FT
Mérida, Yucatán
contacto@certusconsultores.com.mx | Tel: 999 163 1363

═════════════════════════════════════════════════════════════════
`;
  
  return cronograma;
}

async function main() {
  try {
    // Obtener datos de línea de comandos, archivo o defaults
    let datosFormulario = {};
    
    if (process.argv[2]) {
      // Si se pasó JSON por línea de comandos
      try {
        datosFormulario = JSON.parse(process.argv[2]);
      } catch (e) {
        console.error('❌ Error parseando JSON.');
        process.exit(1);
      }
    } else if (fs.existsSync('cliente.json')) {
      // Si existe cliente.json, usa eso
      try {
        const contenido = fs.readFileSync('cliente.json', 'utf8');
        datosFormulario = JSON.parse(contenido);
        console.log('📂 Datos leídos de cliente.json\n');
      } catch (e) {
        console.error('❌ Error leyendo cliente.json');
        process.exit(1);
      }
    } else {
      // Si no hay argumentos ni archivo, usa valores por defecto (para testing)
      datosFormulario = {
        "La Empresa": "Test Inmobiliaria",
        "RFC": "TST123456789",
        "Ciudad": "Mérida"
      };
      console.log('⚠️  Sin datos. Usando valores de prueba...\n');
    }
    
    const empresa = datosFormulario['La Empresa'] || 'Cliente';
    const rfc = datosFormulario['RFC'] || 'RFC';
    const ciudad = datosFormulario['Ciudad'] || 'Mérida';
    
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
    
    // Mapa de reemplazos dinámicos
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
      
      // Reemplaza todos los placeholders
      for (const [placeholder, valor] of Object.entries(reemplazos)) {
        docXml = docXml.split(placeholder).join(valor);
      }
      
      zip.file('word/document.xml', docXml);
      const personalizado = await zip.generateAsync({ type: 'nodebuffer' });
      
      docsPersonalizados[archivo] = personalizado;
      console.log(`  ✓ ${archivo}`);
    }
    
console.log(`\n📅 Generando cronograma...\n`);
    const cronograma = await generarCronograma(empresa, rfc);
    const nombreCronograma = `CRONOGRAMA_${empresa.replace(/\s/g, '_')}_${Date.now()}.txt`;
    fs.writeFileSync(nombreCronograma, cronograma, 'utf8');
    console.log(`✓ Cronograma creado: ${nombreCronograma}\n`);

    console.log(`\n📦 Comprimiendo documentos...\n`);
    const zipFinal = new JSZip();
    for (const [nombre, buffer] of Object.entries(docsPersonalizados)) {
      zipFinal.file(nombre, buffer);
    }
    zipFinal.file(nombreCronograma, fs.readFileSync(nombreCronograma));

    const zipBuffer = await zipFinal.generateAsync({ type: 'nodebuffer' });
    const nombreZip = `CERTUS_PLD_${empresa.replace(/\s/g, '_')}_${Date.now()}.zip`;
    fs.writeFileSync(nombreZip, zipBuffer);
    console.log(`✓ ZIP creado: ${nombreZip}`);
    console.log(`  Tamaño: ${(zipBuffer.length / 1024 / 1024).toFixed(2)} MB\n`);

    console.log(`✅ ¡COMPLETADO!\n`);
    console.log(`📍 Archivos generados:`);
    console.log(`   • ${nombreZip}`);
    console.log(`   • ${nombreCronograma}\n`);
    console.log(`💡 Próxima vez: node generador.js '{"La Empresa":"Nuevo Cliente","RFC":"XYZ123",...}'\n`);
    
  } catch (error) {
    console.error('\n❌ ERROR:', error.message);
    console.error(error);
    process.exit(1);
  }
}

main();
