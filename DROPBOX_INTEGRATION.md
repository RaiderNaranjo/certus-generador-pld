# 🚀 INTEGRACIÓN CON DROPBOX — ZIP → Dropbox → Descarga Manual

## FLUJO FINAL

```
Tally (usuario llena form)
   ↓
Make.com webhook
   ↓
POST /api/generar-documentos (a Railway)
   ↓
server.js (Express):
   1. Crea directorio temporal
   2. Escribe cliente.json
   3. Ejecuta generador.js
   4. Personaliza 14 .docx
   5. Comprime en ZIP
   6. ✅ SUBE A DROPBOX
   ↓
Dropbox: `/Certus_PLD/Clientes/[RFC]_[empresa]_[timestamp].zip`
   ↓
TÚ DESCARGAS: Cuando lo necesites 📥
```

---

## CAMBIOS REALIZADOS

### Archivo: `server-dropbox.js` (NUEVO)
- Basado en `server.js` anterior
- **Agregada integración Dropbox**
- Usa librería `axios` para API de Dropbox
- Sube ZIP automáticamente después de generarlo
- Devuelve información del archivo en Dropbox

---

## INSTALACIÓN Y CONFIGURACIÓN

### PASO 1: Reemplazar server.js

En tu carpeta `certus-railway-v2/`:
- Borra: `server.js` (el anterior)
- Renombra: `server-dropbox.js` → `server.js`

O simplemente copia el contenido de `server-dropbox.js` a `server.js`

### PASO 2: Crear archivo .env

Copia el contenido de `.env.dropbox` a `.env`:

```
NODE_ENV=production
PORT=3000

DROPBOX_TOKEN=sl.u.AGcEbcGVENnYazBe8RISpDEXVgJYUCTprNm5wwVCW-saxE7dKkTwt7gSrMpDicB02G4Z4zeMjUQaxTSFNtBZiA21OvOM8GeXfTGeciUPek6Q2fux4L6WMky9_Iibj2yPJZetPp_9AHGKDNFfIp1PH6GmbIyDZiG2GXDo_EaAxWcE9EAcSEl-2Ja8GWorOOLt1mbmletoOCCDJIyffz34PFeHc21b7kbjQQ92t30E75UyROMKo2tg9MqTscOuAvkIgJgPZ0x-gzBKiv6Bqu7Q7V1hkyblSO-kD3BFbFdSDtnd2mU__mhWcb4ndWdALP6ZPg4U3QZsfCOcDr_Y4NxpKFNuMAPk3WSQgZG1aJtx1NNItYUW36VglA3p3JevJVhAMt-5Mi556DIR4zLZAFW3N-A8LEdjQ467m8h0FQvZu95Vmjo1qCo6BRboZWnssscUfh0Vo7K_9j6lPJaFvho255Zhu-ORJCEfbNwg1SHCuIkUws7RwASPHg1spjtDIKMahj1p4lRx8Jz4EuQNcu-w2uHJRTsnQ6TOrtBvc3KoiMqUE223JVAJejtR9KpZ8RAT26NH8Etz_8xvR7Rho0IQDrkEMC60jkZLxJMt9bn6krB8jRE2j5-lSzMEg9XygB7QiNDpfCPCbVg2SfzUvPTEiLux5-OklXZVDX42_Mz3kBh7jYLX9ADkPpvZcxPcl6TxfoqIGmXXb8svWwm0p2n8oXWGW-tpzBQBAfAxuu45pivKoihZf5dbmDuxYFSM3oac83KEFUNFKqFlOGGnITeVC7rGxvMFOkzqpfRKtKnl3aKq_xQ1H1TXMfasrFoaMBzg0gGKDt-JaEYL48qvDD5y0YeY3cPsy2WGdnB8wg_-zywBrFxForDk8q9LW21rmf0ZrdydzPVjPWIYUOo9vNIRp9IghltN2fqIROgx1-GSkL-5RhqWrNqYstY2czXuvZhBsvuWuXKQdnYUVZYyv4S1BlpMZ7L6uU5JY_XesEMRaNGzr_eH9vXer_DgE5NHnIbkYWGCKgevM0_AN66DKbkChVUm2oNZWw4thtnkvmFjJ6gNPnssUy8TWwjtPd4Us-qnr56kGINy_tuiX51Fy80xGIj-pbHXai2L3VxuxsU58Pqws6_ykw_jmNJ4gWS0d0OBtPEQeYFjyWxkX4OcHPcR0zWY7OPxpQJtZz8G8ZelyL2xTxW401US8iSUTJbBa1_8xSj_sdgejboEgARpW31pLcWHyaemWI4clM5sSNND1WbDTMS1wvdr0t3qqCwV-ljNjuYbFaEmlbqdHn-F_DGV_OHiTojyUFeafvWk4oIV7ReEhw

DROPBOX_FOLDER=/Certus_PLD/Clientes
```

### PASO 3: Actualizar package.json

Asegúrate de tener `axios` en las dependencias:

```json
{
  "dependencies": {
    "express": "^4.18.2",
    "cors": "^2.8.5",
    "jszip": "^3.10.1",
    "archiver": "^6.0.0",
    "dotenv": "^16.3.1",
    "axios": "^1.6.0"
  }
}
```

Si no está, agrega:
```bash
npm install axios
```

---

## PROBAR LOCALMENTE

```bash
cd certus-railway-v2

npm install

npm start
```

Deberías ver:
```
🚀 CERTUS PLD — SERVIDOR DE AUTOMATIZACIÓN v2.1
☁️  Almacenamiento:
   • Dropbox: Configurado
   • Carpeta: /Certus_PLD/Clientes
```

Test con curl:
```bash
curl -X POST http://localhost:3000/api/generar-documentos \
  -H "Content-Type: application/json" \
  -d '{
    "cliente": {
      "La_Empresa": "Test Inmobiliaria",
      "RFC": "TST120101ABC",
      "Domicilio": "Calle Test 123",
      "Ciudad": "Mérida",
      "Actividades": "Fracción V",
      "UMA": "117.31",
      "RepresentantePLD": "Test User",
      "Email": "test@example.com"
    }
  }'
```

### Respuesta esperada:

```json
{
  "success": true,
  "requestId": "abc12345",
  "cliente": {
    "empresa": "Test Inmobiliaria",
    "rfc": "TST120101ABC",
    "ciudad": "Mérida",
    "email": "test@example.com"
  },
  "documentos": {
    "generados": 14,
    "tamaño_mb": "2.45",
    "archivo_local": "CERTUS_PLD_Test_Inmobiliaria_1234567890.zip"
  },
  "dropbox": {
    "subido": true,
    "ruta": "/Certus_PLD/Clientes/TST120101ABC_Test_Inmobiliaria_1234567890.zip",
    "nombre": "TST120101ABC_Test_Inmobiliaria_1234567890.zip",
    "id": "id:a1234567890abcdef",
    "tamaño_bytes": 2563456,
    "carpeta": "/Certus_PLD/Clientes"
  },
  "timestamp": "2026-05-19T14:32:15.000Z",
  "mensaje": "✅ Kit PLD generado y subido a Dropbox exitosamente"
}
```

Verifica en tu Dropbox:
- Abre Dropbox
- Ve a `/Certus_PLD/Clientes/`
- ✅ Deberías ver el ZIP ahí

---

## SUBIR A GITHUB Y RAILWAY

```bash
git add .
git commit -m "🚀 Integración Dropbox para almacenamiento de ZIPs"
git push -u origin main
```

Railway detecta cambios → auto-deploy

---

## ACTUALIZAR MAKE.COM

**URL:** `https://tu-proyecto.railway.app/api/generar-documentos`

**Payload:**
```json
{
  "cliente": {
    "La_Empresa": "{{nombre_empresa}}",
    "RFC": "{{rfc}}",
    "Domicilio": "{{domicilio}}",
    "Ciudad": "{{ciudad}}",
    "Actividades": "{{actividades}}",
    "UMA": "117.31",
    "RepresentantePLD": "{{representante_pld}}",
    "Email": "{{email_cliente}}"
  }
}
```

**Response:** JSON con información del archivo en Dropbox

---

## DESCARGAR ZIP DESDE DROPBOX

1. Abre Dropbox en tu navegador o app
2. Ve a `/Certus_PLD/Clientes/`
3. Busca el ZIP por RFC del cliente
4. Click derecho → Descargar
5. ✅ ZIP en tu disco duro

---

## LOGS EN RAILWAY

Verás algo como:

```
📨 [abc12345] 2026-05-19T14:32:15.000Z
   Solicitud recibida de 203.0.113.42

   📋 Datos del cliente:
      • Empresa: Inmobiliaria XYZ
      • RFC: IXY121010XYZ
      • Ciudad: Mérida
      • Email: contacto@empresa.com

   ✅ Archivo cliente.json creado

   ⚙️  Ejecutando generador de documentos...
      [logs del generador...]

   ✅ Documentos generados correctamente

   📦 ZIP generado localmente:
      • Nombre: CERTUS_PLD_Inmobiliaria_XYZ_1234567890.zip
      • Tamaño: 2.45 MB

   📤 Subiendo a Dropbox...
      Ruta: /Certus_PLD/Clientes/IXY121010XYZ_Inmobiliaria_XYZ_1234567890.zip

   ✅ Archivo subido a Dropbox exitosamente
      ID: id:a1234567890abcdef
      Tamaño: 2.45 MB

   📊 Resumen:
      • Estado: EXITOSO
      • Documentos: 14
      • Tamaño: 2.45 MB
      • Ubicación: /Certus_PLD/Clientes/IXY121010XYZ_Inmobiliaria_XYZ_1234567890.zip

   ✅ [abc12345] Proceso completado
```

---

## ESTRUCTURA FINAL

```
certus-railway-v2/
├── server.js                    (← server-dropbox.js renombrado)
├── generador.js                 (tu código)
├── plantillas/                  (14 documentos)
├── .env                         (con DROPBOX_TOKEN)
├── .env.dropbox                 (template)
├── .gitignore                   (EXCLUYE .env)
├── package.json                 (con axios)
└── [otros archivos]
```

---

## DROPBOX FOLDER STRUCTURE

```
Dropbox Root
└── Certus_PLD/
    └── Clientes/
        ├── IXY121010XYZ_Inmobiliaria_XYZ_1234567890.zip
        ├── TST120101ABC_Test_Inmobiliaria_1234567890.zip
        └── [más ZIPs...]
```

---

## PREGUNTAS FRECUENTES

### ¿El ZIP se genera completamente antes de subir?
Sí. Primero genera todo en Railway, luego sube a Dropbox.

### ¿Cuánto tarda el proceso?
5-15 segundos total (depende del tamaño de documentos).

### ¿Qué pasa si Dropbox falla?
Devuelve error 500 con detalles de la falla. Puedes reintentar.

### ¿Se borran los archivos temporales?
Sí, a los 30 segundos de completar se borra el directorio temporal en Railway.

### ¿Cuánto espacio usa Dropbox?
Cada ZIP es ~2-3 MB. Con 10 clientes/mes = ~250 MB/mes. Dropbox gratis da 2GB.

### ¿Cómo sé que subió correctamente?
Verifica en la respuesta JSON el campo `"dropbox": { "subido": true, ... }`

---

## CONTACTO

Email: contacto@certusconsultores.com.mx

---

*Certus PLD — Generador integrado con Dropbox* ☁️
