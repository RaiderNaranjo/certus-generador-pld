const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');
const { v4: uuidv4 } = require('uuid');
require('dotenv').config();

// ─────────────────────────────────────────────────────────────────────────────
// CONFIGURACIÓN
// ─────────────────────────────────────────────────────────────────────────────
const app = express();
const PORT = process.env.PORT || 3000;
const NODE_ENV = process.env.NODE_ENV || 'development';

// Middleware
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ limit: '10mb', extended: true }));

// Directorios
const TEMP_DIR = path.join(__dirname, '.temp');
const OUTPUT_DIR = path.join(__dirname, 'outputs');

// Crear directorios si no existen
[TEMP_DIR, OUTPUT_DIR].forEach(dir => {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// HEALTHCHECK
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
    versión: '2.2.0',
    ambiente: NODE_ENV,
    endpoints: [
      'GET  /health — Verificar estado',
      'GET  /api/info — Esta información',
      'POST /api/generar-documentos — Generar kit personalizado'
    ],
    generador: 'Integrado (generador.js)',
    documentos: '14 documentos personalizables (D0–D13)',
    almacenamiento: 'Carpeta local (/outputs)',
    autor: 'José Luis Naranjo Bobadilla',
    contacto: 'contacto@certusconsultores.com.mx'
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// ENDPOINT PRINCIPAL: GENERAR DOCUMENTOS Y GUARDAR LOCALMENTE
// ─────────────────────────────────────────────────────────────────────────────
app.post('/api/generar-documentos', async (req, res) => {
  const requestId = uuidv4().slice(0, 8);
  const timestamp = new Date().toISOString();
  
  console.log(`\n${'═'.repeat(80)}`);
  console.log(`📨 [${requestId}] ${timestamp}`);
  console.log(`   Solicitud recibida de ${req.ip}`);

  try {
    // Extraer datos del payload
    const { cliente } = req.body;
    
    if (!cliente) {
      console.log(`❌ [${requestId}] Payload inválido: falta objeto 'cliente'`);
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
      Email = 'cliente@empresa.com'
    } = cliente;

    console.log(`\n   📋 Datos del cliente:`);
    console.log(`      • Empresa: ${La_Empresa}`);
    console.log(`      • RFC: ${RFC}`);
    console.log(`      • Ciudad: ${Ciudad}`);
    console.log(`      • Email: ${Email}`);

    // Crear directorio temporal para este cliente
    const clientDir = path.join(TEMP_DIR, `${RFC}_${requestId}`);
    fs.mkdirSync(clientDir, { recursive: true });

    // Crear archivo cliente.json para el generador
    const datosCliente = {
      'La Empresa': La_Empresa,
      'RFC': RFC,
      'Ciudad': Ciudad,
      'Domicilio': Domicilio,
      'Actividades': Actividades,
      'UMA': UMA,
      'RepresentantePLD': RepresentantePLD,
      'Email': Email
    };

    const clientJsonPath = path.join(clientDir, 'cliente.json');
    fs.writeFileSync(clientJsonPath, JSON.stringify(datosCliente, null, 2));
    console.log(`\n   ✅ Archivo cliente.json creado`);

    // Ejecutar generador.js
    console.log(`\n   ⚙️  Ejecutando generador de documentos...`);

    const generarResult = await new Promise((resolve, reject) => {
      const proceso = spawn('node', [path.join(__dirname, 'generador.js')], {
        cwd: clientDir,
        stdio: ['pipe', 'pipe', 'pipe']
      });

      let stdout = '';
      let stderr = '';

      proceso.stdout.on('data', (data) => {
        stdout += data.toString();
        console.log(`      ${data.toString().trim()}`);
      });

      proceso.stderr.on('data', (data) => {
        stderr += data.toString();
        console.log(`      ⚠️  ${data.toString().trim()}`);
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

    console.log(`\n   ✅ Documentos generados correctamente`);

    // Buscar el ZIP generado
    const files = fs.readdirSync(clientDir);
    const zipFile = files.find(f => f.endsWith('.zip'));

    if (!zipFile) {
      throw new Error('No se generó archivo ZIP');
    }

    const zipPath = path.join(clientDir, zipFile);
    const zipStats = fs.statSync(zipPath);

    console.log(`\n   📦 ZIP generado localmente:`);
    console.log(`      • Nombre: ${zipFile}`);
    console.log(`      • Tamaño: ${(zipStats.size / 1024 / 1024).toFixed(2)} MB`);

    // ✅ GUARDAR EN CARPETA /outputs
    const nombreFinal = `${RFC}_${La_Empresa.replace(/\s/g, '_')}_${Date.now()}.zip`;
    const rutaFinal = path.join(OUTPUT_DIR, nombreFinal);
    
    console.log(`\n   💾 Guardando en carpeta local...`);
    console.log(`      Ruta: /outputs/${nombreFinal}`);

    fs.copyFileSync(zipPath, rutaFinal);

    console.log(`   ✅ Archivo guardado exitosamente`);
    console.log(`      Tamaño: ${(zipStats.size / 1024 / 1024).toFixed(2)} MB`);

    // Preparar respuesta
    const respuesta = {
      success: true,
      requestId: requestId,
      cliente: {
        empresa: La_Empresa,
        rfc: RFC,
        ciudad: Ciudad,
        email: Email
      },
      documentos: {
        generados: 14,
        tamaño_mb: (zipStats.size / 1024 / 1024).toFixed(2),
        archivo_local: zipFile
      },
      almacenamiento: {
        tipo: 'Carpeta local',
        ruta: `/outputs/${nombreFinal}`,
        nombre: nombreFinal,
        tamaño_bytes: zipStats.size,
descargar: `GET /descargar/${nombreFinal}`
      },
      timestamp: new Date().toISOString(),
      mensaje: `✅ Kit PLD generado y guardado exitosamente en carpeta local`
    };

    console.log(`\n   📊 Resumen:`);
    console.log(`      • Estado: EXITOSO`);
    console.log(`      • Documentos: ${respuesta.documentos.generados}`);
    console.log(`      • Tamaño: ${respuesta.documentos.tamaño_mb} MB`);
    console.log(`      • Ubicación: /outputs/${nombreFinal}`);
    console.log(`\n   ✅ [${requestId}] Proceso completado`);
    console.log(`${'═'.repeat(80)}\n`);

    res.json(respuesta);

    // Limpiar directorio temporal después de 30 segundos
    setTimeout(() => {
      try {
        fs.rmSync(clientDir, { recursive: true, force: true });
        console.log(`   🗑️  Directorio temporal ${clientDir} eliminado`);
      } catch (e) {
        console.log(`   ⚠️  No se pudo limpiar directorio: ${e.message}`);
      }
    }, 30000);

  } catch (error) {
    console.log(`\n   ❌ [${requestId}] Error: ${error.message}`);
    console.log(`${'═'.repeat(80)}\n`);

    res.status(500).json({
      error: 'Error al generar documentos',
      message: error.message,
      requestId
    });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// ENDPOINT: DESCARGAR ZIP
// ─────────────────────────────────────────────────────────────────────────────
app.get('/descargar/:archivo', (req, res) => {
  const nombreArchivo = req.params.archivo;
  const rutaArchivo = path.join(OUTPUT_DIR, nombreArchivo);

  // Seguridad: validar que el archivo existe y está en /outputs
  if (!fs.existsSync(rutaArchivo) || !rutaArchivo.startsWith(OUTPUT_DIR)) {
    console.log(`⚠️  Intento de acceso a archivo no válido: ${nombreArchivo}`);
    return res.status(404).json({ error: 'Archivo no encontrado' });
  }

  console.log(`📥 Descargando: ${nombreArchivo}`);
  res.download(rutaArchivo);
});

// ─────────────────────────────────────────────────────────────────────────────
// ENDPOINT: LISTAR ARCHIVOS GENERADOS
// ─────────────────────────────────────────────────────────────────────────────
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
          descarga: `/descargar/${f}`
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

// ─────────────────────────────────────────────────────────────────────────────
// 404 HANDLER
// ─────────────────────────────────────────────────────────────────────────────
app.use((req, res) => {
  console.log(`⚠️  404: ${req.method} ${req.path}`);
  res.status(404).json({
    error: 'Ruta no encontrada',
    path: req.path,
    sugerencia: 'Ver GET /api/info para endpoints disponibles'
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// ERROR HANDLER
// ─────────────────────────────────────────────────────────────────────────────
app.use((err, req, res, next) => {
  console.error(`❌ Error no manejado:`, err);
  res.status(500).json({
    error: 'Error interno del servidor',
    message: NODE_ENV === 'development' ? err.message : 'Error desconocido'
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// INICIAR SERVIDOR
// ─────────────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\n${'═'.repeat(80)}`);
  console.log(`🚀 CERTUS PLD — SERVIDOR DE AUTOMATIZACIÓN v2.2`);
  console.log(`${'═'.repeat(80)}`);
  console.log(`📍 Puerto: ${PORT}`);
  console.log(`🌍 Ambiente: ${NODE_ENV}`);
  console.log(`⏰ Iniciado: ${new Date().toISOString()}`);
  console.log(`\n📡 Endpoints disponibles:`);
  console.log(`   • GET  http://localhost:${PORT}/health`);
  console.log(`   • GET  http://localhost:${PORT}/api/info`);
  console.log(`   • POST http://localhost:${PORT}/api/generar-documentos`);
  console.log(`   • GET  http://localhost:${PORT}/api/archivos`);
  console.log(`   • GET  http://localhost:${PORT}/descargar/:archivo`);
  console.log(`\n💾 Almacenamiento: Carpeta local (/outputs)`);
  console.log(`🔗 Generador: Integrado (generador.js v1.0)`);
  console.log(`📄 Documentos: 14 (D0–D13) personalizables`);
  console.log(`${'═'.repeat(80)}\n`);
});

// Graceful shutdown
process.on('SIGTERM', () => {
  console.log('\n⚠️  SIGTERM recibido. Cerrando servidor...');
  process.exit(0);
});
