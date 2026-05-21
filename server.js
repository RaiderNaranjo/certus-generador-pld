const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');
const crypto = require('crypto');
const https = require('https');
const { Readable } = require('stream');
const jwt = require('jsonwebtoken');

const uuidv4 = () => crypto.randomUUID();

try {
  require('dotenv').config();
} catch (e) {
  console.log('⚠️ .env no encontrado');
}

// CONFIGURACIÓN
const app = express();
const PORT = process.env.PORT || 3000;
const NODE_ENV = process.env.NODE_ENV || 'development';
const DRIVE_FOLDER_ID = process.env.GOOGLE_DRIVE_FOLDER_ID;

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ limit: '10mb', extended: true }));

const TEMP_DIR = path.join(__dirname, '.temp');
const OUTPUT_DIR = path.join(__dirname, 'outputs');

[TEMP_DIR, OUTPUT_DIR].forEach(dir => {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// GOOGLE DRIVE API REST (sin googleapis)
// ─────────────────────────────────────────────────────────────────────────────

async function getGoogleAccessToken() {
  const serviceAccount = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
  
  const payload = {
    iss: serviceAccount.client_email,
    scope: 'https://www.googleapis.com/auth/drive.file',
    aud: 'https://oauth2.googleapis.com/token',
    exp: Math.floor(Date.now() / 1000) + 3600,
    iat: Math.floor(Date.now() / 1000)
  };

  const token = jwt.sign(payload, serviceAccount.private_key, { algorithm: 'RS256' });

  return new Promise((resolve, reject) => {
    const postData = JSON.stringify({
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: token
    });

    const options = {
      hostname: 'oauth2.googleapis.com',
      port: 443,
      path: '/token',
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(postData)
      }
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          resolve(parsed.access_token);
        } catch (e) {
          reject(new Error('Error parsing token response'));
        }
      });
    });

    req.on('error', reject);
    req.write(postData);
    req.end();
  });
}

async function uploadToDrive(zipPath, nombreArchivo) {
  console.log(`  ☁️  Subiendo a Google Drive: ${nombreArchivo}`);

  const accessToken = await getGoogleAccessToken();
  const fileStream = fs.createReadStream(zipPath);
  const fileSize = fs.statSync(zipPath).size;

  return new Promise((resolve, reject) => {
    const boundary = '===============7330845974216740156==';
    const body = [];

    body.push(`--${boundary}`);
    body.push('Content-Disposition: form-data; name="metadata"');
    body.push('Content-Type: application/json\r');
    body.push('');
    body.push(JSON.stringify({
      name: nombreArchivo,
      parents: [DRIVE_FOLDER_ID],
      mimeType: 'application/zip'
    }));
    body.push(`\r--${boundary}`);
    body.push('Content-Disposition: form-data; name="file"; filename="' + nombreArchivo + '"');
    body.push('Content-Type: application/zip\r');
    body.push('');

    const startBytes = Buffer.from(body.join('\r\n') + '\r\n');
    const endBytes = Buffer.from(`\r\n--${boundary}--`);

    const options = {
      hostname: 'www.googleapis.com',
      port: 443,
      path: '/upload/drive/v3/files?uploadType=multipart',
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': `multipart/related; boundary="${boundary}"`,
        'Content-Length': startBytes.length + fileSize + endBytes.length
      }
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const fileData = JSON.parse(data);
          const fileId = fileData.id;
          const downloadUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;

          console.log(`  ✅ Subido a Drive — ID: ${fileId}`);

          // Hacer público (sin esperar)
          makeFilePublic(fileId, accessToken).catch(e => 
            console.log(`  ⚠️ No se pudo hacer público: ${e.message}`)
          );

          resolve({
            fileId,
            nombre: fileData.name,
            downloadUrl,
            viewUrl: `https://drive.google.com/file/d/${fileId}/view`
          });
        } catch (e) {
          reject(new Error(`Error en respuesta Drive: ${e.message}`));
        }
      });
    });

    req.on('error', reject);

    req.write(startBytes);
    fileStream.pipe(req);
    fileStream.on('end', () => req.write(endBytes));
  });
}

async function makeFilePublic(fileId, accessToken) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify({
      role: 'reader',
      type: 'anyone'
    });

    const options = {
      hostname: 'www.googleapis.com',
      port: 443,
      path: `/drive/v3/files/${fileId}/permissions`,
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(body)
      }
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve(JSON.parse(data)));
    });

    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// ENDPOINTS
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
    versión: '3.1.0',
    ambiente: NODE_ENV,
    almacenamiento: 'Google Drive',
    documentos: '14 (D0–D13)',
    autor: 'José Luis Naranjo Bobadilla'
  });
});

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
        message: 'Se requiere objeto "cliente"',
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

    console.log(`  📋 ${La_Empresa} (${RFC})`);

    const clientDir = path.join(TEMP_DIR, `${RFC}_${requestId}`);
    fs.mkdirSync(clientDir, { recursive: true });

    fs.writeFileSync(
      path.join(clientDir, 'cliente.json'),
      JSON.stringify({
        'La Empresa': La_Empresa,
        'RFC': RFC,
        'Ciudad': Ciudad,
        'Domicilio': Domicilio,
        'Actividades': Actividades,
        'UMA': UMA,
        'RepresentantePLD': RepresentantePLD,
        'NombreRepresentanteLegal': NombreRepresentanteLegal,
        'Email': Email
      }, null, 2)
    );

    console.log(`  ⚙️  Ejecutando generador...`);
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

    const files = fs.readdirSync(clientDir);
    const zipFile = files.find(f => f.endsWith('.zip'));

    if (!zipFile) throw new Error('No se generó ZIP');

    const zipPath = path.join(clientDir, zipFile);
    const zipStats = fs.statSync(zipPath);
    const nombreFinal = `${RFC}_${La_Empresa.replace(/\s+/g, '_')}_${Date.now()}.zip`;

    console.log(`  📦 ZIP: ${(zipStats.size / 1024 / 1024).toFixed(2)} MB`);

    let driveInfo = null;
    let downloadUrl = null;

    if (DRIVE_FOLDER_ID && process.env.GOOGLE_SERVICE_ACCOUNT) {
      try {
        driveInfo = await uploadToDrive(zipPath, nombreFinal);
        downloadUrl = driveInfo.downloadUrl;
      } catch (driveError) {
        console.log(`  ⚠️ Error en Drive: ${driveError.message}`);
        const rutaFinal = path.join(OUTPUT_DIR, nombreFinal);
        fs.copyFileSync(zipPath, rutaFinal);
        downloadUrl = `/descargar/${nombreFinal}`;
      }
    } else {
      const rutaFinal = path.join(OUTPUT_DIR, nombreFinal);
      fs.copyFileSync(zipPath, rutaFinal);
      downloadUrl = `/descargar/${nombreFinal}`;
    }

    const respuesta = {
      success: true,
      requestId,
      cliente: { empresa: La_Empresa, rfc: RFC, ciudad: Ciudad, email: Email },
      documentos: { generados: 14, tamaño_mb: parseFloat((zipStats.size / 1024 / 1024).toFixed(2)), nombre: nombreFinal },
      almacenamiento: driveInfo ? {
        tipo: 'Google Drive',
        fileId: driveInfo.fileId,
        nombre: driveInfo.nombre,
        downloadUrl: driveInfo.downloadUrl,
        viewUrl: driveInfo.viewUrl
      } : {
        tipo: 'Local',
        downloadUrl
      },
      downloadUrl,
      timestamp: new Date().toISOString(),
      mensaje: `✅ Kit PLD generado y subido a Google Drive`
    };

    console.log(`  ✅ [${requestId}] Completado`);
    console.log(`${'═'.repeat(80)}\n`);

    res.json(respuesta);

    setTimeout(() => {
      try {
        fs.rmSync(clientDir, { recursive: true, force: true });
      } catch (e) {}
    }, 60000);

  } catch (error) {
    console.log(`  ❌ [${requestId}] ${error.message}`);
    console.log(`${'═'.repeat(80)}\n`);
    res.status(500).json({
      error: 'Error al generar documentos',
      message: error.message,
      requestId
    });
  }
});

app.get('/descargar/:archivo', (req, res) => {
  const nombreArchivo = req.params.archivo;
  const rutaArchivo = path.join(OUTPUT_DIR, nombreArchivo);

  if (!fs.existsSync(rutaArchivo) || !rutaArchivo.startsWith(OUTPUT_DIR)) {
    return res.status(404).json({ error: 'Archivo no encontrado' });
  }

  res.download(rutaArchivo);
});

app.use((req, res) => {
  res.status(404).json({ error: 'Ruta no encontrada' });
});

app.listen(PORT, () => {
  console.log(`\n${'═'.repeat(80)}`);
  console.log(`🚀 CERTUS PLD v3.1 (Google Drive Direct API)`);
  console.log(`${'═'.repeat(80)}`);
  console.log(`📍 Puerto: ${PORT}`);
  console.log(`☁️  Drive: ${DRIVE_FOLDER_ID ? 'configurado' : 'NO CONFIGURADO'}`);
  console.log(`${'═'.repeat(80)}\n`);
});

process.on('SIGTERM', () => {
  console.log('\n⚠️ SIGTERM recibido. Cerrando...');
  process.exit(0);
});
