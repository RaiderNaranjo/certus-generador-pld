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
  allowedHeaders: ['Content-Type', 'x-api-key']
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

// ===== AUTENTICACIÓN (API_KEY en variables de Railway) =====
const auth = (req, res, next) => {
  const key = process.env.API_KEY;
  if (!key) return next();
  if (req.get('x-api-key') !== key) {
    return res.status(401).json({ success: false, error: 'No autorizado' });
  }
  next();
};

const limpiar = (t) => String(t).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Za-z0-9_-]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 60);

// ===== ENDPOINT PRINCIPAL =====
app.post('/api/generar-documentos', auth, async (req, res) => {
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
    const tempDir = path.join(__dirname, '.temp', `${limpiar(cliente.RFC)}_${requestId}`);
    const outputDir = path.join(__dirname, 'outputs');
    
    fs.mkdirSync(tempDir, { recursive: true });
    fs.mkdirSync(outputDir, { recursive: true });

    // Escribir cliente.json en el directorio temporal
    const clienteJsonPath = path.join(tempDir, 'cliente.json');
    fs.writeFileSync(clienteJsonPath, JSON.stringify({
      'La Empresa': cliente.La_Empresa,
      'RFC': cliente.RFC,
      'Tipo': (cliente.Tipo || req.body.tipo || 'inmobiliaria'),
      'TipoPersona': (cliente.TipoPersona || req.body.tipoPersona || ''),
      'Ciudad': cliente.Ciudad || 'Mérida',
      'Domicilio': cliente.Domicilio || '',
      'Actividades': cliente.Actividades || '',
      'Email': cliente.Email || '',
      'UMA': cliente.UMA || '117.31',
      'RepresentantePLD': cliente.RepresentantePLD || '',
      'NombreRepresentanteLegal': cliente.NombreRepresentanteLegal || '',
      'Incisos': cliente.Incisos || '',
      'IncisosTexto': cliente.IncisosTexto || '',
      'IncisosNoPrestados': cliente.IncisosNoPrestados || '',
      'NumeroNotaria': cliente.NumeroNotaria || '',
      'NumeroCorreduria': cliente.NumeroCorreduria || '',
      'Plaza': cliente.Plaza || '',
      'Profesion': cliente.Profesion || '',
      'CedulaProfesional': cliente.CedulaProfesional || '',
      'ActuaEnRepresentacion': cliente.ActuaEnRepresentacion || '',
      'LineasExcluidas': cliente.LineasExcluidas || '',
      'EncargadoOperativo': cliente.EncargadoOperativo || '',
      'Telefono': cliente.TelefonoRL || '',
      'FechaFirma': cliente.FechaFirma || req.body.fechaFirma || '',
      'Consultor': cliente.Consultor || req.body.consultor || ''
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
    const nombreFinal = `${limpiar(cliente.RFC)}_${limpiar(cliente.La_Empresa)}_${Date.now()}.zip`;
    const rutaFinal = path.join(outputDir, nombreFinal);
    fs.copyFileSync(zipPath, rutaFinal);

    console.log(`💾 ZIP guardado: ${nombreFinal}`);

    const totalDocs = (/Personalizando (\d+) documentos/.exec(generarResult.stdout) || [])[1] || 0;

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
        generados: Number(totalDocs) + 1,
        total: Number(totalDocs) + 1,
        nombre: nombreFinal,
        tamaño_mb: (zipStats.size / 1024 / 1024).toFixed(2)
      },
      downloadUrl: downloadUrl,
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

// ===== PRE-DIAGNÓSTICO PARA EL CLIENTE (PDF) =====
// Body: { cliente, perfil, respuestas }  — lo envía Make tal cual lo armó el intake (campo prediagnosticoJSON)
const prediagnostico = require('./prediagnostico');
app.post('/api/prediagnostico', auth, async (req, res) => {
  try {
    let datos = req.body || {};
    if (typeof datos.prediagnosticoJSON === 'string') datos = JSON.parse(datos.prediagnosticoJSON);
    const cliente = datos.cliente || {};
    if (!cliente.La_Empresa || !cliente.RFC || !datos.respuestas) {
      return res.status(400).json({ success: false, error: 'Faltan datos: cliente.La_Empresa, cliente.RFC y respuestas' });
    }
    const { resultado, buffer } = await prediagnostico.generarPDF(datos);
    const nombre = `Prediagnostico_PLD_${limpiar(cliente.RFC)}_${Date.now()}.pdf`;
    fs.writeFileSync(path.join(__dirname, 'outputs', nombre), buffer);
    const baseURL = process.env.BASE_URL || 'https://certus-generador-pld-production.up.railway.app';
    console.log(`📄 Pre-diagnóstico ${resultado.folio}: ${resultado.indice}/100 (${resultado.nivel})`);
    res.json({
      success: true,
      resultado: {
        folio: resultado.folio, indice: resultado.indice, nivel: resultado.nivel,
        inherente: resultado.inherente, noSe: resultado.noSe,
        bloques: resultado.bloques, pendientes: resultado.pendientes
      },
      correo: prediagnostico.correoCliente(datos, resultado),
      pdfNombre: nombre,
      pdfUrl: `${baseURL}/descargar/${nombre}`,
      pdfBase64: buffer.toString('base64')
    });
  } catch (err) {
    console.error('Pre-diagnóstico:', err);
    res.status(500).json({ success: false, error: String(err.message || err) });
  }
});

// ===== CONSULTA DE LISTAS (KYC de Certus) =====
const listas = require('./listas');
app.post('/api/consultar-listas', auth, async (req, res) => {
  try {
    const r = await listas.consultar(req.body || {});
    res.json(Object.assign({ success: true }, r));
  } catch (err) {
    console.error('Consulta de listas:', err);
    res.status(500).json({ success: false, error: String(err.message || err) });
  }
});
app.get('/api/listas/estado', auth, (req, res) => {
  res.json(Object.keys(listas.FUENTES).map(id => Object.assign({ id, nombre: listas.FUENTES[id].nombre }, listas.estado[id] || {}, { cargando: undefined })));
});

// ===== DESCARGAR =====
app.get('/descargar/:archivo', (req, res) => {
  const archivo = path.basename(req.params.archivo);
  const rutaArchivo = path.join(__dirname, 'outputs', archivo);

  if (!fs.existsSync(rutaArchivo)) {
    return res.status(404).json({ error: 'Archivo no encontrado' });
  }

  res.download(rutaArchivo);
});

// ===== INICIAR =====
app.listen(PORT, () => {
  setTimeout(() => listas.asegurar().then(() => console.log('📋 Listas cargadas:', JSON.stringify(Object.keys(listas.estado).map(k => k + '=' + (listas.estado[k].registros || 0) + (listas.estado[k].error ? ' (' + listas.estado[k].error + ')' : ''))))).catch(e => console.error('Listas:', e.message)), 2000);
  console.log(`\n${'═'.repeat(60)}`);
  console.log(`🚀 CERTUS PLD — SERVIDOR DE AUTOMATIZACIÓN`);
  console.log(`${'═'.repeat(60)}`);
  console.log(`📍 Puerto: ${PORT}`);
  console.log(`🔌 CORS: ✅ Habilitado para todos los orígenes`);
  console.log(`📁 Generador: generador.js (Node.js)`);
  console.log(`📄 Plantillas: ${path.join(__dirname, 'plantillas')}`);
  console.log(`${'═'.repeat(60)}\n`);
});
