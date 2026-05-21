const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');
const crypto = require('crypto');

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
const BASE_URL = process.env.BASE_URL || 'https://certus-generador-pld-production.up.railway.app';

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

// Servir archivos estáticos de outputs
app.use('/descargar', express.static(OUTPUT_DIR));

// ─────────────────────────────────────────────────────────────────────────────
// ENDPOINTS
// ─────────────────────────────────────────────────────────────────────────────

app.get('/health', (req, res) => {
  res.status(200).json({
    status: 'online',
    server: 'Certus PLD Automation',
    environment: NODE_ENV,
    timestamp: new Date().toISOString(),
    uptime: Math.floor(process.uptime())
  });
});

app.get('/api/info', (req, res) => {
  res.json({
    nombre: 'Certus PLD — Servidor de Automatización',
    versión: '4.0.0',
    ambiente: NODE_ENV,
    almacenamiento: 'Local (Railway)',
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
    console.log(`  📧 ${Email}`);

    const clientDir = path.join(TEMP_DIR, `${RFC}_${requestId}`);
    fs.mkdirSync(clientDir, { recursive: true });

    // Crear cliente.json
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

    // Ejecutar generador
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
        else reject(new Error(`Generador código ${code}`));
      });
      proceso.on('error', err => reject(new Error(`Error: ${err.message}`)));
    });

    // Buscar ZIP
    const files = fs.readdirSync(clientDir);
    const zipFile = files.find(f => f.endsWith('.zip'));

    if (!zipFile) throw new Error('No se generó ZIP');

    const zipPath = path.join(clientDir, zipFile);
    const zipStats = fs.statSync(zipPath);
    const nombreFinal = `${RFC}_${La_Empresa.replace(/\s+/g, '_')}_${Date.now()}.zip`;

    console.log(`  📦 ZIP: ${(zipStats.size / 1024 / 1024).toFixed(2)} MB`);

    // Guardar en outputs
    const rutaFinal = path.join(OUTPUT_DIR, nombreFinal);
    fs.copyFileSync(zipPath, rutaFinal);

    const downloadUrl = `${BASE_URL}/descargar/${nombreFinal}`;

    console.log(`  💾 Guardado en: /outputs/${nombreFinal}`);
    console.log(`  🔗 Link: ${downloadUrl}`);

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
      almacenamiento: {
        tipo: 'Local (Railway)',
        nombre: nombreFinal,
        downloadUrl: downloadUrl
      },
      downloadUrl,
      timestamp: new Date().toISOString(),
      mensaje: `✅ Kit PLD generado y listo para descargar`
    };

    console.log(`  ✅ [${requestId}] Completado`);
    console.log(`${'═'.repeat(80)}\n`);

    res.json(respuesta);

    // Limpiar temp después de 60 segundos
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

  console.log(`  📥 Descargando: ${nombreArchivo}`);
  res.download(rutaArchivo);
});

app.get('/api/archivos', (req, res) => {
  try {
    const archivos = fs.readdirSync(OUTPUT_DIR)
      .filter(f => f.endsWith('.zip'))
      .map(f => {
        const stats = fs.statSync(path.join(OUTPUT_DIR, f));
        return {
          nombre: f,
          tamaño_mb: (stats.size / 1024 / 1024).toFixed(2),
          fecha: new Date(stats.mtimeMs).toISOString(),
          downloadUrl: `${BASE_URL}/descargar/${f}`
        };
      })
      .sort((a, b) => new Date(b.fecha) - new Date(a.fecha));

    res.json({
      total: archivos.length,
      archivos: archivos
    });
  } catch (error) {
    res.status(500).json({
      error: 'Error listando archivos',
      message: error.message
    });
  }
});

app.use((req, res) => {
  res.status(404).json({ error: 'Ruta no encontrada', path: req.path });
});

app.listen(PORT, () => {
  console.log(`\n${'═'.repeat(80)}`);
  console.log(`🚀 CERTUS PLD v4.0 (Simple & Reliable)`);
  console.log(`${'═'.repeat(80)}`);
  console.log(`📍 Puerto: ${PORT}`);
  console.log(`💾 Almacenamiento: Local (Railway /outputs)`);
  console.log(`${'═'.repeat(80)}\n`);
});

process.on('SIGTERM', () => {
  console.log('\n⚠️ SIGTERM. Cerrando...');
  process.exit(0);
});
