const express = require('express');
const cors = require('cors');
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// CORS
app.use(cors({
  origin: '*',
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type']
}));

app.use(express.json({ limit: '50mb' }));

// Crear carpeta outputs
if (!fs.existsSync('outputs')) {
  fs.mkdirSync('outputs', { recursive: true });
}

// ===== HEALTH CHECK =====
app.get('/health', (req, res) => {
  res.json({ 
    status: 'OK',
    service: 'Certus PLD Generator',
    timestamp: new Date().toISOString()
  });
});

// ===== ENDPOINT PRINCIPAL =====
app.post('/api/generar-documentos', async (req, res) => {
  const requestId = Date.now().toString().slice(-8);
  
  try {
    const { cliente } = req.body;

    if (!cliente || !cliente.La_Empresa || !cliente.RFC) {
      return res.status(400).json({ 
        success: false, 
        error: 'Faltan datos: La_Empresa y RFC' 
      });
    }

    console.log(`\n🔨 [${requestId}] Generando documentos para ${cliente.La_Empresa}`);

    // Crear directorio temporal
    const timestamp = Date.now();
    const tempDir = path.join(__dirname, '.temp', `${cliente.RFC}_${requestId}`);
    const outputDir = path.join(__dirname, 'outputs');
    
    fs.mkdirSync(tempDir, { recursive: true });
    fs.mkdirSync(outputDir, { recursive: true });

    // Escribir cliente.json en el directorio temporal
    const clienteJsonPath = path.join(tempDir, 'cliente.json');
    fs.writeFileSync(clienteJsonPath, JSON.stringify({
      'La Empresa': cliente.La_Empresa,
      'RFC': cliente.RFC,
      'Ciudad': cliente.Ciudad || 'Mérida',
      'Domicilio': cliente.Domicilio || '',
      'Actividades': cliente.Actividades || '',
      'Email': cliente.Email || '',
      'UMA': cliente.UMA || '117.31',
      'RepresentantePLD': cliente.RepresentantePLD || '',
      'NombreRepresentanteLegal': cliente.NombreRepresentanteLegal || ''
    }));

    // Ejecutar generador.js
    const generarResult = await new Promise((resolve, reject) => {
      const proceso = spawn('node', [path.join(__dirname, 'generador.js')], {
        cwd: tempDir,
        env: { ...process.env, PLANTILLAS_PATH: path.join(__dirname, 'plantillas') }
      });

      let stdout = '';
      let stderr = '';

      proceso.stdout.on('data', (data) => {
        stdout += data.toString();
        console.log(`  ${data.toString().trim()}`);
      });

      proceso.stderr.on('data', (data) => {
        stderr += data.toString();
        console.error(`  ⚠️  ${data.toString().trim()}`);
      });

      proceso.on('close', (code) => {
        if (code === 0) {
          resolve({ success: true, stdout, stderr });
        } else {
          reject(new Error(`Generador salió con código ${code}: ${stderr}`));
        }
      });

      proceso.on('error', (err) => {
        reject(new Error(`Error ejecutando generador: ${err.message}`));
      });
    });

    console.log(`✅ Documentos generados correctamente`);

    // Buscar el ZIP generado
    const files = fs.readdirSync(tempDir);
    const zipFile = files.find(f => f.endsWith('.zip'));

    if (!zipFile) {
      throw new Error('No se generó archivo ZIP');
    }

    const zipPath = path.join(tempDir, zipFile);
    const zipStats = fs.statSync(zipPath);

    // Guardar en outputs
    const nombreFinal = `${cliente.RFC}_${cliente.La_Empresa.replace(/\s+/g, '_')}_${Date.now()}.zip`;
    const rutaFinal = path.join(outputDir, nombreFinal);
    fs.copyFileSync(zipPath, rutaFinal);

    console.log(`💾 ZIP guardado: ${nombreFinal}`);

    // Leer como base64
    const zipBuffer = fs.readFileSync(rutaFinal);
    const zipBase64 = zipBuffer.toString('base64');

    const baseURL = process.env.BASE_URL || 'https://certus-generador-pld-production.up.railway.app';
    const downloadUrl = `${baseURL}/descargar/${nombreFinal}`;

    res.json({
      success: true,
      requestId: requestId,
      cliente: {
        empresa: cliente.La_Empresa,
        rfc: cliente.RFC,
        ciudad: cliente.Ciudad || 'Mérida',
        email: cliente.Email || ''
      },
      documentos: {
        generados: 17,
        total: 17,
        nombre: nombreFinal,
        tamaño_mb: (zipStats.size / 1024 / 1024).toFixed(2)
      },
      downloadUrl: downloadUrl,
      zipBase64: zipBase64,
      zipNombre: nombreFinal
    });

    // Limpiar temporal después de 30 segundos
    setTimeout(() => {
      try {
        fs.rmSync(tempDir, { recursive: true, force: true });
      } catch (e) {
        console.log(`⚠️  No se pudo limpiar ${tempDir}`);
      }
    }, 30000);

  } catch (error) {
    console.error(`❌ [${requestId}] Error: ${error.message}`);
    res.status(500).json({
      success: false,
      error: error.message,
      requestId: requestId
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
  console.log(`\n${'═'.repeat(60)}`);
  console.log(`🚀 CERTUS PLD — SERVIDOR DE AUTOMATIZACIÓN`);
  console.log(`${'═'.repeat(60)}`);
  console.log(`📍 Puerto: ${PORT}`);
  console.log(`🔌 CORS: ✅ Habilitado para todos los orígenes`);
  console.log(`📁 Generador: generador.js (Node.js)`);
  console.log(`📄 Plantillas: ${path.join(__dirname, 'plantillas')}`);
  console.log(`${'═'.repeat(60)}\n`);
});
