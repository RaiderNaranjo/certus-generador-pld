const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');
const crypto = require('crypto');
const { google } = require('googleapis');

const uuidv4 = () => crypto.randomUUID();

try {
  require('dotenv').config();
} catch (e) {
  console.log('⚠️ .env no encontrado, usando valores por defecto');
}

// ─────────────────────────────────────────────────────────────────────────────
// CONFIGURACIÓN
// ─────────────────────────────────────────────────────────────────────────────
const app = express();
const PORT = process.env.PORT || 3000;
const NODE_ENV = process.env.NODE_ENV || 'development';
const DRIVE_FOLDER_ID = process.env.GOOGLE_DRIVE_FOLDER_ID;

// Middleware
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ limit: '10mb', extended: true }));

// Directorios
const TEMP_DIR = path.join(__dirname, '.temp');
const OUTPUT_DIR = path.join(__dirname, 'outputs');

[TEMP_DIR, OUTPUT_DIR].forEach(dir => {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// GOOGLE DRIVE CLIENT
// ─────────────────────────────────────────────────────────────────────────────
function getDriveClient() {
  const serviceAccount = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
  const auth = new google.auth.GoogleAuth({
    credentials: serviceAccount,
    scopes: ['https://www.googleapis.com/auth/drive.file']
  });
  return google.drive({ version: 'v3', auth });
}

async function subirADrive(zipPath, nombreArchivo) {
  const drive = getDriveClient();

  console.log(`  ☁️  Subiendo a Google Drive: ${nombreArchivo}`);

  const response = await drive.files.create({
    requestBody: {
      name: nombreArchivo,
      parents: [DRIVE_FOLDER_ID],
    },
    media: {
      mimeType: 'application/zip',
      body: fs.createReadStream(zipPath),
    },
    fields: 'id, name, webContentLink, webViewLink'
  });

  const fileId = response.data.id;

  // Hacer el archivo público para descarga directa
  await drive.permissions.create({
    fileId: fileId,
    requestBody: {
      role: 'reader',
      type: 'anyone'
    }
  });

  // Obtener link de descarga directa
  const fileData = await drive.files.get({
    fileId: fileId,
    fields: 'id, name, webContentLink, webViewLink'
  });

  const downloadUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;

  console.log(`  ✅ Subido exitosamente a Drive`);
  console.log(`  🔗 Link: ${downloadUrl}`);

  return {
    fileId,
    nombre: fileData.data.name,
    downloadUrl,
    viewUrl: fileData.data.webViewLink
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// HEALTHCHECK
// ─────────────────────────────────────────────────────────────────────────────
app.get('/health', (req, res) => {
  res.status(200).json({
    status: 'online',
    server: 'Certus PLD Automation',
    environment: NODE_ENV,
    timestamp: new Date().toISOString(),
    uptime: Math.floor(process.uptime()),
    drive: DRIVE_FOLDER_ID ? 'configurado' : 'no configurado'
  });
});

app.get('/api/info', (req, res) => {
  res.json({
    nombre: 'Certus PLD — Servidor de Automatización',
    versión: '3.0.0',
    ambiente: NODE_ENV,
    endpoints: [
      'GET /health — Verificar estado',
      'GET /api/info — Esta información',
      'POST /api/generar-documentos — Generar kit personalizado'
    ],
    almacenamiento: 'Google Drive',
    documentos: '14 documentos personalizables (D0–D13)',
    autor: 'José Luis Naranjo Bobadilla',
    contacto: 'contacto@certusconsultores.com.mx'
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// ENDPOINT PRINCIPAL: GENERAR DOCUMENTOS Y SUBIR A GOOGLE DRIVE
// ─────────────────────────────────────────────────────────────────────────────
app.post('/api/generar-documentos', async (req, res) => {
  const requestId = uuidv4().slice(0, 8);
  const timestamp = new Date().toISOString();

  console.log(`\n${'═'.repeat(80)}`);
  console.log(`📨 [${requestId}] ${timestamp}`);

  try {
    const { cliente } = req.body;

    if (!cliente) {
      return res.status(400).json({
        error: 'Payload inválido',
        message: 'Se requiere un objeto "cliente" con los datos',
        requestId
      });
    }

    const {
      La_Empresa = 'SIN_ESPECIFICAR',
      RFC = 'XXX000000ABC',
      Domicilio = 'Domicilio no especificado',
      Ciudad = 'Mérida',
      Actividades = 'No especificadas',
      UMA = '117.31',
      RepresentantePLD = 'No asignado',
      NombreRepresentanteLegal = 'No especificado',
      Email = 'cliente@empresa.com'
    } = cliente;

    console.log(`\n  📋 Cliente: ${La_Empresa} (${RFC})`);
    console.log(`  📧 Email: ${Email}`);

    // Crear directorio temporal
    const clientDir = path.join(TEMP_DIR, `${RFC}_${requestId}`);
    fs.mkdirSync(clientDir, { recursive: true });

    // Crear cliente.json
    const datosCliente = {
      'La Empresa': La_Empresa,
      'RFC': RFC,
      'Ciudad': Ciudad,
      'Domicilio': Domicilio,
      'Actividades': Actividades,
      'UMA': UMA,
      'RepresentantePLD': RepresentantePLD,
      'NombreRepresentanteLegal': NombreRepresentanteLegal,
      'Email': Email
    };

    fs.writeFileSync(
      path.join(clientDir, 'cliente.json'),
      JSON.stringify(datosCliente, null, 2)
    );

    // Ejecutar generador.js
    console.log(`\n  ⚙️  Ejecutando generador...`);
    await new Promise((resolve, reject) => {
      const proceso = spawn('node', [path.join(__dirname, 'generador.js')], {
        cwd: clientDir,
        stdio: ['pipe', 'pipe', 'pipe']
      });

      proceso.stdout.on('data', d => console.log(`  ${d.toString().trim()}`));
      proceso.stderr.on('data', d => console.log(`  ⚠️ ${d.toString().trim()}`));

      proceso.on('close', code => {
        if (code === 0) resolve();
        else reject(new Error(`Generador salió con código ${code}`));
      });

      proceso.on('error', err => reject(new Error(`Error: ${err.message}`)));
    });

    // Buscar ZIP generado
    const files = fs.readdirSync(clientDir);
    const zipFile = files.find(f => f.endsWith('.zip'));

    if (!zipFile) throw new Error('No se generó archivo ZIP');

    const zipPath = path.join(clientDir, zipFile);
    const zipStats = fs.statSync(zipPath);
    const nombreFinal = `${RFC}_${La_Empresa.replace(/\s+/g, '_')}_${Date.now()}.zip`;

    console.log(`\n  📦 ZIP: ${zipFile} (${(zipStats.size / 1024 / 1024).toFixed(2)} MB)`);

    // Subir a Google Drive
    let driveInfo = null;
    let downloadUrl = null;

    if (DRIVE_FOLDER_ID && process.env.GOOGLE_SERVICE_ACCOUNT) {
      try {
        driveInfo = await subirADrive(zipPath, nombreFinal);
        downloadUrl = driveInfo.downloadUrl;
      } catch (driveError) {
        console.log(`  ⚠️ Error subiendo a Drive: ${driveError.message}`);
        // Fallback: guardar localmente
        const rutaFinal = path.join(OUTPUT_DIR, nombreFinal);
        fs.copyFileSync(zipPath, rutaFinal);
        downloadUrl = `/descargar/${nombreFinal}`;
      }
    } else {
      // Sin Drive configurado: guardar localmente
      const rutaFinal = path.join(OUTPUT_DIR, nombreFinal);
      fs.copyFileSync(zipPath, rutaFinal);
      downloadUrl = `/descargar/${nombreFinal}`;
    }

    // Respuesta
    const respuesta = {
      success: true,
      requestId,
      cliente: {
        empresa: La_Empresa,
        rfc: RFC,
        ciudad: Ciudad,
        email: Email
      },
      documentos: {
        generados: 14,
        tamaño_mb: parseFloat((zipStats.size / 1024 / 1024).toFixed(2)),
        nombre: nombreFinal
      },
      almacenamiento: driveInfo ? {
        tipo: 'Google Drive',
        fileId: driveInfo.fileId,
        nombre: driveInfo.nombre,
        downloadUrl: driveInfo.downloadUrl,
        viewUrl: driveInfo.viewUrl
      } : {
        tipo: 'Carpeta local',
        downloadUrl
      },
      downloadUrl,
      timestamp: new Date().toISOString(),
      mensaje: `✅ Kit PLD generado y subido a Google Drive exitosamente`
    };

    console.log(`\n  ✅ [${requestId}] Completado — Link: ${downloadUrl}`);
    console.log(`${'═'.repeat(80)}\n`);

    res.json(respuesta);

    // Limpiar temp después de 60 segundos
    setTimeout(() => {
      try {
        fs.rmSync(clientDir, { recursive: true, force: true });
      } catch (e) {}
    }, 60000);

  } catch (error) {
    console.log(`\n  ❌ [${requestId}] Error: ${error.message}`);
    console.log(`${'═'.repeat(80)}\n`);
    res.status(500).json({
      error: 'Error al generar documentos',
      message: error.message,
      requestId
    });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// ENDPOINT DESCARGA LOCAL (fallback)
// ─────────────────────────────────────────────────────────────────────────────
app.get('/descargar/:archivo', (req, res) => {
  const nombreArchivo = req.params.archivo;
  const rutaArchivo = path.join(OUTPUT_DIR, nombreArchivo);

  if (!fs.existsSync(rutaArchivo) || !rutaArchivo.startsWith(OUTPUT_DIR)) {
    return res.status(404).json({ error: 'Archivo no encontrado' });
  }

  res.download(rutaArchivo);
});

// ─────────────────────────────────────────────────────────────────────────────
// 404 + ERROR HANDLERS
// ─────────────────────────────────────────────────────────────────────────────
app.use((req, res) => {
  res.status(404).json({ error: 'Ruta no encontrada', path: req.path });
});

app.use((err, req, res, next) => {
  console.error('❌ Error no manejado:', err);
  res.status(500).json({ error: 'Error interno del servidor' });
});

// ─────────────────────────────────────────────────────────────────────────────
// INICIAR
// ─────────────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\n${'═'.repeat(80)}`);
  console.log(`🚀 CERTUS PLD — SERVIDOR v3.0 (Google Drive)`);
  console.log(`${'═'.repeat(80)}`);
  console.log(`📍 Puerto: ${PORT}`);
  console.log(`☁️  Drive Folder: ${DRIVE_FOLDER_ID || 'NO CONFIGURADO'}`);
  console.log(`${'═'.repeat(80)}\n`);
});

process.on('SIGTERM', () => {
  console.log('\n⚠️ SIGTERM recibido. Cerrando servidor...');
  process.exit(0);
});
