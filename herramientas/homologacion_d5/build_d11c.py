"""D11C — Evaluación de Comprensión, homologada para los cuatro kits (12 preguntas, 24 puntos, aprobatorio 17).

Uso: python3 build_d11c.py <dir_contenedor> <dir_salida>
Toma las preguntas vigentes de cada kit (src/d11c/<KIT>.docx), agrega las del Módulo 7 donde faltan y arma la clave.
"""
import os, sys, re
sys.path.insert(0, os.path.dirname(__file__))
import build_d5 as B
from build_d5 import Ctx, p, tabla, portada, pie, escribir_par, Par, aplicar_globales
from docxlib import instalar_estilos, q, espaciar_tras_tablas, leer_par
from docx import Document
import canon
import build_d11 as D11

def texto(el):
    return ''.join(t.text or '' for t in el.iter(q('w:t')))

def leer_preguntas(path):
    doc = Document(path)
    pregs = []; cur = None
    for el in doc.element.body:
        if el.tag != q('w:p'):
            if cur is not None and el.tag == q('w:tbl') and 'Clave de respuestas' in texto(el):
                cur = None
            continue
        t = texto(el).strip()
        if t.startswith('Clave de respuestas'):
            cur = None; continue
        m = re.match(r'Pregunta (\d+)\s*[—-]\s*(.*)', t)
        if m:
            cur = dict(n=int(m.group(1)), titulo=m.group(2), enunciado=[], opciones=[]); pregs.append(cur); continue
        if cur is None or not t or t.startswith(('CERTUS CONSULTORES', 'Pregunta abierta', 'Preguntas de opción')):
            continue
        if t.startswith('[Espacio') or t.startswith('Respuesta:') or set(t) <= set('_ '):
            continue
        if re.match(r'\[ \] [A-D]\)', t):
            cur['opciones'].append(leer_par(el))
        else:
            cur['enunciado'].append(leer_par(el))
    return pregs

M7 = [dict(titulo='Grado de Riesgo (Módulo 7)',
           enunciado='¿Cada cuánto debe reevaluarse, como mínimo, el Grado de Riesgo de un cliente?',
           opciones=['Una sola vez, al integrar el expediente',
                     'Al menos cada seis meses, y con mayor frecuencia cuanto mayor sea el Grado de Riesgo',
                     'Cada cinco años, al actualizar la identificación',
                     'Solo cuando la autoridad lo requiera']),
      dict(titulo='Auditoría anual (Módulo 7)',
           enunciado='¿Cuál de las siguientes afirmaciones sobre la auditoría anual PLD/FT es CORRECTA conforme al Acuerdo 115/2026?',
           opciones=['Es una práctica recomendada, no una obligación legal',
                     'Es obligatoria desde el ejercicio 2026',
                     'Es obligatoria; el primer periodo auditado es 2028 y, si el Riesgo del sujeto obligado es bajo o medio, el dictamen puede emitirlo un área de auditoría o control interno independiente',
                     'Solo aplica a quienes tienen más de 50 empleados'])]
K10 = ('B', 'art. 23 Bis 1 RCG; D6, sección IX')
K11 = ('C', 'art. 18, fr. XI, de la Ley; arts. 42, 44 y 45 RCG; Transitorio Octavo del Acuerdo 115/2026')

CLAVES = {
 'INM': {1: ('C', 'Fracción V Bis (reforma de julio de 2025), sin requisito de habitualidad'),
         2: ('C', '8,025 UMA × $117.31 = $941,413 MXN'),
         3: ('B', 'Efectivo parcial, saldo indefinido y origen no explicado son señales de alerta; el valor ($800,000) es inferior al umbral del art. 32, que se mide contra el valor del inmueble'),
         4: ('C', 'art. 18, fr. I, de la Ley; art. 12 RCG: identificación previa a la operación'),
         5: ('B', 'Los estados de cuenta son el soporte estándar para acreditar ahorros (D7, Sección C)'),
         6: ('B', 'El familiar directo de una PEP es persona asimilada (art. 23 Quáter RCG; D8)'),
         7: ('B', '$300,000 supera el umbral de identificación (1,605 UMA = $188,283) pero no el de Aviso (3,210 UMA = $376,565)'),
         8: ('C', 'Escalar sin comunicarlo al cliente: confidencialidad (art. 50 de la Ley; art. 31 RCG; D9, sección V)'),
         9: ('B', 'El informe en ceros es obligatorio aunque no haya operaciones (art. 25 RCG)'),
         10: ('C', 'Riesgo alto por criterio del Programa, que extiende el art. 23 Bis 4 RCG al Beneficiario Controlador no residente, con independencia del puntaje EBR'),
         11: ('C', 'art. 18, fr. XI, de la Ley; arts. 42, 44 y 45 RCG; Transitorio Octavo del Acuerdo 115/2026'),
         12: ('Abierta', 'Al menos tres señales identificadas; escalar al responsable; identificar al Beneficiario Controlador; no aceptar efectivo (art. 32, fr. I) y aplicar medidas reforzadas o no operar')},
 'NOT': {1: ('B', 'art. 17, fr. XII, Apartado A, incisos b), c) y e), de la Ley'),
         2: ('B', 'Inciso a): 8,000 UMA × $117.31 = $938,480 MXN'),
         3: ('C', 'art. 32, fr. I, de la Ley: el límite se mide contra el valor del inmueble; documentar en el D9-NOT'),
         4: ('C', 'art. 18, fr. III, de la Ley; art. 23 Quinquies RCG'),
         5: ('B', 'Soporte documental del origen de los recursos (D7-NOT, Sección C)'),
         6: ('B', 'La calidad de PEP es de la persona, compre o venda (art. 23 Quáter RCG)'),
         7: ('C', 'Inciso e): siempre objeto de Aviso; evaluar el Riesgo con el D6-NOT'),
         8: ('B', 'art. 18, fr. VI, de la Ley; art. 26 Bis RCG'),
         9: ('C', 'art. 58 de la Ley'),
         10: K10, 11: K11,
         12: ('Abierta', 'La persona física cumple personal y directamente (art. 20 de la Ley); solo las personas morales y las figuras jurídicas designan Representante Encargado de Cumplimiento. El Encargado Operativo apoya en la ejecución sin trasladar la responsabilidad (D2-NOT, sección III)')},
 'COR': {1: ('B', 'art. 17, fr. XII, Apartado B, incisos b), c) y d), de la Ley'),
         2: ('C', 'Inciso a): 8,025 UMA × $117.31 = $941,413 MXN'),
         3: ('B', 'art. 32, fr. VI, de la Ley: 3,210 UMA = $376,565 MXN'),
         4: ('C', 'art. 18, fr. III, de la Ley; art. 23 Quinquies RCG'),
         5: ('B', 'art. 19, primer párrafo, RCG: los incisos del Apartado B no se acumulan; Aviso de 24 horas si hay fraccionamiento artificial'),
         6: ('B', 'La calidad de PEP es de la persona, compre o venda (art. 23 Quáter RCG)'),
         7: ('C', 'Inciso d): siempre objeto de Aviso; evaluar el Riesgo con el D6-COR'),
         8: ('B', 'art. 18, fr. VI, de la Ley; art. 26 Bis RCG'),
         9: ('C', 'art. 58 de la Ley (y cancelación de la habilitación por reincidencia, art. 57)'),
         10: K10, 11: K11,
         12: ('Abierta', 'La persona física cumple personal y directamente (art. 20 de la Ley); solo las personas morales y las figuras jurídicas designan Representante Encargado de Cumplimiento. El Encargado Operativo apoya en la ejecución sin trasladar la responsabilidad (D2-COR, sección III)')},
 'PROF': {1: ('B', 'Sin umbral monetario: se activa al preparar el servicio (D1-PROF, sección 2.2)'),
          2: ('B', 'La fracción XI no fue modificada por el Decreto DOF 16-07-2025'),
          3: ('B', 'Señal de alerta; abstención si no se proporciona la información (art. 21 de la Ley)'),
          4: ('C', 'art. 18, fr. III, de la Ley; art. 23 Quinquies RCG'),
          5: ('B', 'La calidad de PEP es de la persona, sin importar el contexto (art. 23 Quáter RCG; D8-PROF)'),
          6: ('B', 'Identificación desde la preparación; Aviso solo con representación (D1-PROF, sección 2.2)'),
          7: ('B', 'art. 17, fr. XI, último párrafo, de la Ley; D1-PROF, sección 2.2'),
          8: ('B', 'art. 18, fr. VI, de la Ley; art. 26 Bis RCG'),
          9: ('B', 'D1-PROF, sección IV; art. 22 de la Ley'),
          10: K10, 11: K11,
          12: ('Abierta', 'Inciso d); identificación obligatoria desde la preparación; documentar la alerta y evaluar la abstención (art. 21 de la Ley) antes de continuar')},
}

ARREGLOS = [(r'matriz ponderada de D6', 'Metodología EBR del D6'),
            (r'\(independiente del RC\)', '(independiente del Representante Encargado de Cumplimiento)'),
            (r'\bRC\b', 'Representante Encargado de Cumplimiento'),
            (r'el SAT notifica a la autoridad estatal', 'la Secretaría informa a la autoridad estatal'),
            (r'inciso a-e\b', 'inciso a) a e)'), (r'incisos a-e\b', 'incisos a) a e)')]

def limpiar(par):
    for a, b in ARREGLOS:
        par.replace(a, b)
    aplicar_globales(par)
    return par

def construir(kit):
    ctx = Ctx(kit); d = D11.S[kit]; s = d['suf']
    pregs = leer_preguntas(os.path.join(os.path.dirname(__file__), '..', 'src', 'd11c', '%s.docx' % kit))
    mc = [x for x in pregs if x['opciones']]
    abierta = [x for x in pregs if not x['opciones']][-1]
    if len(mc) == 9:
        for m in M7:
            mc.append(dict(titulo=m['titulo'], enunciado=[Par.from_text(m['enunciado'])],
                           opciones=[Par.from_text('[ ] %s) %s' % ('ABCD'[i], o)) for i, o in enumerate(m['opciones'])]))
    assert len(mc) == 11, (kit, len(mc))
    portada(ctx, codigo='D11C', titulo='Evaluación de Comprensión',
            indice='Índice: I. Opción múltiple (11 preguntas) │ II. Pregunta abierta │ III. Clave de respuestas (instructor)')
    grupos = '   '.join('[ ] G%d' % i for i in range(1, len(d['grupos']) + 1))
    tabla(ctx, [('**Nombre del participante:**', ''), ('**Cargo:**', ''), ('**Grupo (D11%s):**' % s, grupos),
                ('**Fecha:**', '[dd/mm/aaaa]'), ('**Módulos evaluados:**', '[ ] 1   [ ] 2   [ ] 3   [ ] 4   [ ] 5   [ ] 6   [ ] 7'),
                ('**Calificación:**', '_____ / 24 puntos   │   [ ] APROBADO (17 o más)   [ ] REPROBADO (menos de 17)')],
          [3200, 6300], encabezado=False)
    p(ctx, '**Instrucciones:** marque la respuesta correcta en cada pregunta de opción múltiple (2 puntos cada una) y '
           'conteste la pregunta abierta (2 puntos). Puntaje máximo: 24. Mínimo aprobatorio: 17 (70%%). Quien no apruebe '
           'sigue las medidas del %s, sección 12.1.' % ('D5' + s))
    ctx.body.append(escribir_par(Par.from_text('Sección I: Preguntas de Opción Múltiple (11 × 2 = 22 puntos)'), 'CertusCapitulo'))
    for i, x in enumerate(mc, 1):
        ctx.body.append(escribir_par(Par.from_text('Pregunta %d — %s' % (i, x['titulo'])), 'CertusSeccion'))
        for pp in x['enunciado']:
            ctx.body.append(escribir_par(limpiar(pp), 'CertusTexto'))
        for pp in x['opciones']:
            ctx.body.append(escribir_par(limpiar(pp), 'CertusTexto'))
    ctx.body.append(escribir_par(Par.from_text('Sección II: Pregunta Abierta (2 puntos)'), 'CertusCapitulo'))
    ctx.body.append(escribir_par(Par.from_text('Pregunta 12 — %s' % abierta['titulo']), 'CertusSeccion'))
    for pp in abierta['enunciado']:
        ctx.body.append(escribir_par(limpiar(pp), 'CertusTexto'))
    tabla(ctx, [('**Respuesta del participante:**',), ('',), ('',), ('',), ('',)], [9500], encabezado=False)
    ctx.body.append(escribir_par(Par.from_text('\f'), 'CertusTexto'))
    ctx.body.append(escribir_par(Par.from_text('Sección III: Clave de Respuestas — Uso Exclusivo del Instructor'), 'CertusCapitulo'))
    clave = CLAVES[kit]
    tabla(ctx, [('Pregunta', 'Respuesta', 'Fundamento')] + [(str(n), clave[n][0], clave[n][1].replace('D6,', 'D6%s,' % s)) for n in range(1, 13)],
          [1100, 1300, 7100])
    pie(ctx)
    return ctx

NOMBRES = {'PROF': 'profesionistas/D11C-PROF_Evaluacion_Comprension_v2.docx', 'INM': 'inmobiliaria/D11C_Evaluacion_Comprension_CertusPLD_v5.docx',
           'NOT': 'notarial/D11C-NOT_Evaluacion_Comprension_v3.docx', 'COR': 'corredores/D11C-COR_Evaluacion_Comprension_v2.docx'}

if __name__ == '__main__':
    B.SRC, OUT = sys.argv[1], sys.argv[2]
    for k in ('PROF', 'INM', 'NOT', 'COR'):
        ctx = construir(k)
        doc = Document(os.path.join(B.SRC, 'PROF_pr12.docx'))
        instalar_estilos(doc)
        body = doc.element.body; sect = body.find(q('w:sectPr'))
        for el in list(body):
            if el is not sect: body.remove(el)
        espaciar_tras_tablas(ctx.body)
        for el in ctx.body: sect.addprevious(el)
        cp = doc.core_properties
        cp.title = 'D11C%s Evaluación de Comprensión — %s' % (canon.KITS[k]['suf'], canon.KITS[k]['variante'])
        cp.author = 'Certus Consultores · Espinosa-Salazar Abogados Corporativos'
        cp.subject = 'Programa Certus PLD — edición homologada noviembre 2026'
        dest = os.path.join(OUT, NOMBRES[k]); os.makedirs(os.path.dirname(dest), exist_ok=True)
        from docxlib import numero_de_pagina; numero_de_pagina(doc)
        doc.save(dest); print(k, dest)
