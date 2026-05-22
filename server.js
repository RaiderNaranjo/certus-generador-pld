const express = require('express');
const cors = require('cors');
const { Document, Packer, Paragraph, TextRun } = require('docx');
const archiver = require('archiver');
const fs = require('fs');
const path = require('path');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// ===== CONFIGURACIÓN =====
app.use(cors()); // Permitir CORS desde cualquier dominio
app.use(express.json({ limit: '50mb' }));
app.use(express.static('outputs'));

// ===== PLANTILLAS DE DOCUMENTOS =====
const crearDocumento = (nombreDoc, cliente) => {
  const contenido = {
    'D0-Política-PLD': `POLÍTICA DE CUMPLIMIENTO PLD/FT\n\nEmpresa: ${cliente.La_Empresa}\nRFC: ${cliente.RFC}\nFecha: ${new Date().toLocaleDateString('es-MX')}`,
    'D1-Procedimiento-KYC': `PROCEDIMIENTO KYC\n\nRepresentante PLD: ${cliente.RepresentantePLD}\nActividades: ${cliente.Actividades}`,
    'D2-Matriz-Riesgo': `MATRIZ DE RIESGO\n\nUMA Vigente: $${cliente.UMA} MXN\nDomicilio: ${cliente.Domicilio}, ${cliente.Ciudad}`,
    'D3-Expediente-SPPLD': `EXPEDIENTE SPPLD\n\nRepresentante Legal: ${cliente.NombreRepresentanteLegal}\nEmail: ${cliente.Email}`,
    'D4-Gap-Analysis': `GAP ANALYSIS DE CUMPLIMIENTO\n\nEmpresa: ${cliente.La_Empresa}\nFecha Análisis: ${new Date().toLocaleDateString('es-MX')}`,
    'D5-Registro-Clientes': `REGISTRO DE CLIENTES\n\nEmpresa: ${cliente.La_Empresa}\nRFC: ${cliente.RFC}`,
    'D6-Señales-Alerta': `SEÑALES DE ALERTA PLD\n\nActividades: ${cliente.Actividades}\nUMA: $${cliente.UMA}`,
    'D7-Procedimiento-Avisos': `PROCEDIMIENTO DE AVISOS SAT\n\nDeadline: 17 días hábiles\nPortal: SPPLD`,
    'D8-Registro-Operaciones': `REGISTRO DE OPERACIONES INUSUALES\n\nEmpresa: ${cliente.La_Empresa}\nFecha: ${new Date().toLocaleDateString('es-MX')}`,
    'D9-Capacitación-Plan': `PLAN DE CAPACITACIÓN\n\nRepresentante: ${cliente.RepresentantePLD}\nDuración: 4 módulos`,
    'D10-Evaluación-Inicial': `EVALUACIÓN INICIAL DE CUMPLIMIENTO\n\nEmpresa: ${cliente.La_Empresa}\nFecha: ${new Date().toLocaleDateString('es-MX')}`,
    'D11-Modulos-Capacitacion': `MÓDULOS DE CAPACITACIÓN\n\n1. Marco Normativo LFPIORPI\n2. Señales de Alerta\n3. KYC\n4. Reporte y Documentación`,
    'D12-Libro-Avisos': `LIBRO DE AVISOS\n\nEmpresa: ${cliente.La_Empresa}\nPeriodo: ${new Date().getFullYear()}`,
    'D13-Auditoria-Interna': `AUDITORÍA INTERNA PLD\n\nEmpresa: ${cliente.La_Empresa}\nFecha: ${new Date().toLocaleDateString('es-MX')}`,
    'D14-Manual-Operativo': `MANUAL OPERATIVO DE CUMPLIMIENTO\n\nEmpresa: ${cliente.La_Empresa}\nVersión: 1.0`
  };

  return new Document({
    sections: [{
      children: [
        new Paragraph({
          text: contenido[nombreDoc] || 'Documento generado',
          style: 'normal'
        })
      ]
    }]
  });
};

// ===== ENDPOINT PRINCIPAL =====
app.post('/api/generar-documentos', async (req, res) => {
  console.log('📨 Solicitud recibida:', req.body);

  try {
    const { cliente } = req.body;

    // Validar datos mínimos
    if (!cliente || !cliente.La_Empresa || !cliente.RFC) {
      return res.status(400).json({
        success: false,
        error: 'Faltan datos requeridos: La_Empresa y RFC'
      });
    }

    // Crear carpeta temporal para este cliente
    const timestamp = Date.now();
    const nombreCarpeta = `${cliente.RFC}_${cliente.La_Empresa.replace(/\s+/g, '_')}_${timestamp}`;
    const rutaCarpeta = path.join(__dirname, 'outputs', nombreCarpeta);

    if (!fs.existsSync(path.join(__dirname, 'outputs'))) {
      fs.mkdirSync(path.join(__dirname, 'outputs'), { recursive: true });
    }
    if (!fs.existsSync(rutaCarpeta)) {
      fs.mkdirSync(rutaCarpeta, { recursive: true });
    }

    console.log(`📁 Carpeta creada: ${rutaCarpeta}`);

    // Generar 14 documentos
    const documentos = [
      'D0-Política-PLD',
      'D1-Procedimiento-KYC',
      'D2-Matriz-Riesgo',
      'D3-Expediente-SPPLD',
      'D4-Gap-Analysis',
      'D5-Registro-Clientes',
      'D6-Señales-Alerta',
      'D7-Procedimiento-Avisos',
      'D8-Registro-Operaciones',
      'D9-Capacitación-Plan',
      'D10-Evaluación-Inicial',
      'D11-Modulos-Capacitacion',
      'D12-Libro-Avisos',
      'D13-Auditoria-Interna',
      'D14-Manual-Operativo'
    ];

    let archivosGenerados = 0;

    for (const doc of documentos) {
      try {
        const documento = crearDocumento(doc, cliente);
        const bytes = await Packer.toBuffer(documento);
        const rutaArchivo = path.join(rutaCarpeta, `${doc}.docx`);
        fs.writeFileSync(rutaArchivo, bytes);
        console.log(`✓ Generado: ${doc}.docx`);
        archivosGenerados++;
      } catch (err) {
        console.error(`✗ Error generando ${doc}:`, err.message);
      }
    }

    // Crear ZIP con todos los documentos
    const nombreZip = `${cliente.RFC}_${cliente.La_Empresa.replace(/\s+/g, '_')}.zip`;
    const rutaZip = path.join(rutaCarpeta, nombreZip);

    const output = fs.createWriteStream(rutaZip);
    const archive = archiver('zip', { zlib: { level: 9 } });

    await new Promise((resolve, reject) => {
      archive.on('error', reject);
      archive.pipe(output);

      for (const doc of documentos) {
        const rutaArchivo = path.join(rutaCarpeta, `${doc}.docx`);
        if (fs.existsSync(rutaArchivo)) {
          archive.file(rutaArchivo, { name: `${doc}.docx` });
        }
      }

      archive.finalize();
      output.on('close', resolve);
    });

    console.log(`📦 ZIP creado: ${nombreZip}`);

    // Respuesta exitosa
    const baseURL = process.env.BASE_URL || `https://${req.get('host')}`;
    const downloadUrl = `${baseURL}/descargar/${nombreZip}`;

    const respuesta = {
      success: true,
      requestId: timestamp,
      cliente: {
        empresa: cliente.La_Empresa,
        rfc: cliente.RFC,
        ciudad: cliente.Ciudad,
        email: cliente.Email
      },
      documentos: {
        generados: archivosGenerados,
        total: documentos.length,
        nombre: nombreZip
      },
      almacenamiento: {
        tipo: 'Local (Railway)',
        nombre: nombreZip,
        downloadUrl: downloadUrl
      },
      downloadUrl: downloadUrl
    };

    console.log('✅ Respuesta exitosa:', respuesta);
    res.json(respuesta);

  } catch (error) {
    console.error('❌ Error en generación:', error.message);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// ===== ENDPOINT DE DESCARGA =====
app.get('/descargar/:archivo', (req, res) => {
  const archivo = req.params.archivo;
  const rutaArchivo = path.join(__dirname, 'outputs', archivo);

  console.log(`📥 Solicitando descarga: ${archivo}`);

  if (!fs.existsSync(rutaArchivo)) {
    console.log(`❌ Archivo no encontrado: ${rutaArchivo}`);
    return res.status(404).json({ error: 'Archivo no encontrado' });
  }

  res.download(rutaArchivo, (err) => {
    if (err) console.error('Error descargando:', err);
  });
});

// ===== HEALTH CHECK =====
app.get('/health', (req, res) => {
  res.json({ status: 'OK', timestamp: new Date().toISOString() });
});

// ===== INICIAR SERVIDOR =====
app.listen(PORT, () => {
  console.log(`\n🚀 Servidor Certus PLD corriendo en puerto ${PORT}`);
  console.log(`📍 API disponible en: http://localhost:${PORT}/api/generar-documentos`);
  console.log(`✅ CORS habilitado — acepta requests desde cualquier dominio\n`);
});
