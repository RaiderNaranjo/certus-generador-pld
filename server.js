const express = require('express');
const cors = require('cors');
const { Document, Packer, Paragraph, TextRun, HeadingLevel } = require('docx');
const archiver = require('archiver');
const fs = require('fs');
const path = require('path');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// CORS SIN RESTRICCIONES
app.use(cors({
  origin: '*',
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Accept']
}));

app.use(express.json({ limit: '50mb' }));
app.use(express.static('outputs'));

// Crear carpeta outputs si no existe
if (!fs.existsSync('outputs')) {
  fs.mkdirSync('outputs', { recursive: true });
}

// ===== CREAR DOCUMENTO =====
const crearDocumento = (nombreDoc, cliente) => {
  const hoy = new Date().toLocaleDateString('es-MX');
  
  const contenidos = {
    'D0-Política-PLD': `POLÍTICA DE CUMPLIMIENTO PLD/FT\n\nEmpresa: ${cliente.La_Empresa}\nRFC: ${cliente.RFC}\nFecha: ${hoy}`,
    'D1-Procedimiento-KYC': `PROCEDIMIENTO KYC\n\nRepresentante PLD: ${cliente.RepresentantePLD}\nActividades: ${cliente.Actividades}`,
    'D2-Matriz-Riesgo': `MATRIZ DE RIESGO\n\nUMA Vigente: $${cliente.UMA} MXN\nDomicilio: ${cliente.Domicilio}, ${cliente.Ciudad}`,
    'D3-Expediente-SPPLD': `EXPEDIENTE SPPLD\n\nRepresentante Legal: ${cliente.NombreRepresentanteLegal}\nEmail: ${cliente.Email}`,
    'D4-Gap-Analysis': `GAP ANALYSIS DE CUMPLIMIENTO\n\nEmpresa: ${cliente.La_Empresa}\nFecha: ${hoy}`,
    'D5-Registro-Clientes': `REGISTRO DE CLIENTES\n\nEmpresa: ${cliente.La_Empresa}\nRFC: ${cliente.RFC}`,
    'D6-Señales-Alerta': `SEÑALES DE ALERTA PLD\n\nActividades: ${cliente.Actividades}\nUMA: $${cliente.UMA}`,
    'D7-Procedimiento-Avisos': `PROCEDIMIENTO DE AVISOS SAT\n\nDeadline: 17 días hábiles\nPortal: SPPLD`,
    'D8-Registro-Operaciones': `REGISTRO DE OPERACIONES INUSUALES\n\nEmpresa: ${cliente.La_Empresa}\nFecha: ${hoy}`,
    'D9-Capacitación-Plan': `PLAN DE CAPACITACIÓN\n\nRepresentante: ${cliente.RepresentantePLD}\nDuración: 4 módulos`,
    'D10-Evaluación-Inicial': `EVALUACIÓN INICIAL DE CUMPLIMIENTO\n\nEmpresa: ${cliente.La_Empresa}\nFecha: ${hoy}`,
    'D11-Modulos-Capacitacion': `MÓDULOS DE CAPACITACIÓN\n\n1. Marco Normativo LFPIORPI\n2. Señales de Alerta\n3. KYC\n4. Reporte y Documentación`,
    'D12-Libro-Avisos': `LIBRO DE AVISOS\n\nEmpresa: ${cliente.La_Empresa}\nPeriodo: ${new Date().getFullYear()}`,
    'D13-Auditoria-Interna': `AUDITORÍA INTERNA PLD\n\nEmpresa: ${cliente.La_Empresa}\nFecha: ${hoy}`,
    'D14-Manual-Operativo': `MANUAL OPERATIVO DE CUMPLIMIENTO\n\nEmpresa: ${cliente.La_Empresa}\nVersión: 1.0`
  };

  const contenido = contenidos[nombreDoc] || 'Documento generado automáticamente';

  return new Document({
    sections: [{
      children: [
        new Paragraph({
          text: contenido,
          style: 'Normal'
        })
      ]
    }]
  });
};

// ===== HEALTH CHECK =====
app.get('/health', (req, res) => {
  res.json({ 
    status: 'OK', 
    timestamp: new Date().toISOString(),
    service: 'Certus PLD Generator'
  });
});

// ===== ENDPOINT PRINCIPAL =====
app.post('/api/generar-documentos', async (req, res) => {
  try {
    const { cliente } = req.body;

    if (!cliente || !cliente.La_Empresa || !cliente.RFC) {
      return res.status(400).json({ 
        success: false, 
        error: 'Faltan datos: La_Empresa y RFC' 
      });
    }

    const timestamp = Date.now();
    const nombreCarpeta = `${cliente.RFC}_${timestamp}`;
    const rutaCarpeta = path.join(__dirname, 'outputs', nombreCarpeta);

    fs.mkdirSync(rutaCarpeta, { recursive: true });

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

    let generados = 0;

    // Generar documentos
    for (const doc of documentos) {
      try {
        const documento = crearDocumento(doc, cliente);
        const bytes = await Packer.toBuffer(documento);
        const rutaArchivo = path.join(rutaCarpeta, `${doc}.docx`);
        fs.writeFileSync(rutaArchivo, bytes);
        generados++;
      } catch (err) {
        console.error(`Error en ${doc}:`, err.message);
      }
    }

    // Crear ZIP
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

   // Respuesta
    const baseURL = process.env.BASE_URL || 'https://certus-generador-pld-production.up.railway.app';
    const downloadUrl = `${baseURL}/descargar/${nombreZip}`;

    // Leer ZIP como base64
    const zipBuffer = fs.readFileSync(zipPath);
    const zipBase64 = zipBuffer.toString('base64');

    res.json({
      success: true,
      requestId: timestamp,
      cliente: {
        empresa: cliente.La_Empresa,
        rfc: cliente.RFC,
        ciudad: cliente.Ciudad || '',
        email: cliente.Email || ''
      },
      documentos: {
        generados: generados,
        total: documentos.length,
        nombre: nombreZip
      },
      downloadUrl: downloadUrl,
      zipBase64: zipBase64,
      zipNombre: nombreZip
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// ===== DESCARGAR =====
app.get('/descargar/:archivo', (req, res) => {
  const archivo = req.params.archivo;
  const rutaArchivo = path.join(__dirname, 'outputs', archivo);

  if (!fs.existsSync(rutaArchivo)) {
    return res.status(404).json({ error: 'Archivo no encontrado' });
  }

  res.download(rutaArchivo);
});

// ===== INICIAR =====
app.listen(PORT, () => {
  console.log(`🚀 Certus PLD corriendo en puerto ${PORT}`);
  console.log(`✅ CORS habilitado - acepta requests desde cualquier origen`);
});
