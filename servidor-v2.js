const express = require('express');
const { spawn } = require('child_process');
const path = require('path');
const app = express();

// Middleware
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Log todas las solicitudes
app.use((req, res, next) => {
  console.log(`\n📨 [${new Date().toLocaleTimeString()}] ${req.method} ${req.path}`);
  console.log('Body:', JSON.stringify(req.body, null, 2));
  console.log('Headers:', req.headers);
  next();
});

// Health check
app.get('/', (req, res) => {
  console.log('✅ Health check OK');
  res.json({ status: 'OK', message: 'Servidor Certus PLD activo' });
});

// Main endpoint
app.post('/generar', (req, res) => {
  console.log('\n🎯 POST /generar recibido');
  
  const { La_Empresa, RFC, Ciudad } = req.body;

  console.log(`\n📋 Datos recibidos:`);
  console.log(`   - La_Empresa: ${La_Empresa}`);
  console.log(`   - RFC: ${RFC}`);
  console.log(`   - Ciudad: ${Ciudad}`);

  // Validar datos
  if (!La_Empresa || !RFC) {
    console.log('❌ Datos incompletos');
    return res.status(400).json({ 
      error: 'Faltan datos: La_Empresa y RFC son requeridos',
      recibido: { La_Empresa, RFC, Ciudad }
    });
  }

  console.log('\n⚙️  Ejecutando generador.js...');

  // Crear archivo cliente.json temporal
  const fs = require('fs');
  const datosCliente = {
    'La Empresa': La_Empresa,
    RFC: RFC,
    Ciudad: Ciudad || 'Mérida'
  };

  const rutaCliente = path.join(__dirname, 'cliente-temp.json');
  fs.writeFileSync(rutaCliente, JSON.stringify(datosCliente, null, 2));
  console.log(`✅ Archivo cliente-temp.json creado`);

  // Ejecutar generador.js
  const proceso = spawn('node', ['generador.js'], {
    cwd: __dirname,
    stdio: ['pipe', 'pipe', 'pipe']
  });

  let stdout = '';
  let stderr = '';

  proceso.stdout.on('data', (data) => {
    stdout += data.toString();
    console.log(`📤 stdout: ${data}`);
  });

  proceso.stderr.on('data', (data) => {
    stderr += data.toString();
    console.log(`❌ stderr: ${data}`);
  });

  proceso.on('close', (code) => {
    console.log(`\n✅ Generador terminó con código: ${code}`);
    
    // Limpiar archivo temporal
    try {
      fs.unlinkSync(rutaCliente);
      console.log('🗑️  Archivo temporal eliminado');
    } catch (e) {
      console.log('⚠️  No se pudo eliminar archivo temporal');
    }

    if (code === 0) {
      console.log('🎉 Generación exitosa');
      res.json({ 
        success: true, 
        cliente: La_Empresa,
        rfc: RFC,
        ciudad: Ciudad,
        mensaje: 'Documentos generados correctamente',
        timestamp: new Date()
      });
    } else {
      console.log('❌ Error en generador');
      res.status(500).json({ 
        success: false,
        error: 'Error al generar documentos',
        codigo: code,
        stderr: stderr
      });
    }
  });

  proceso.on('error', (err) => {
    console.log(`❌ Error al ejecutar proceso: ${err.message}`);
    res.status(500).json({ 
      error: 'Error al ejecutar generador',
      details: err.message
    });
  });
});

// 404 handler
app.use((req, res) => {
  console.log(`⚠️  Ruta no encontrada: ${req.method} ${req.path}`);
  res.status(404).json({ error: 'Ruta no encontrada', path: req.path });
});

// Error handler
app.use((err, req, res, next) => {
  console.log(`❌ Error en servidor: ${err.message}`);
  res.status(500).json({ error: 'Error interno del servidor', details: err.message });
});

// Iniciar servidor
const PORT = 3000;
app.listen(PORT, () => {
  console.log(`\n🚀 SERVIDOR CERTUS PLD INICIADO`);
  console.log(`📍 http://localhost:${PORT}`);
  console.log(`🌐 Endpoint: https://lustfully-ardently-profusely.ngrok-free.dev/generar`);
  console.log(`\n⏳ Esperando solicitudes...\n`);
});
