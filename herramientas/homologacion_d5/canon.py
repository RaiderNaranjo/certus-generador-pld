# Estructura canónica del D5 (Manual de Políticas Internas) — cuatro kits Certus PLD
# Fuente única de títulos, índice y correspondencia con el art. 37 Bis RCG.

KITS = {
    'PROF': dict(suf='-PROF', variante='Servicios Profesionales', sujeto='el profesional',
                 frac='Fracción XI del art. 17 LFPIORPI', pfpm=True),
    'INM':  dict(suf='', variante='Inmobiliarias', sujeto='la empresa',
                 frac='Fracciones V, V Bis y XV del art. 17 LFPIORPI', pfpm=True),
    'NOT':  dict(suf='-NOT', variante='Notarías', sujeto='la Notaría',
                 frac='Fracción XII, Apartado A, del art. 17 LFPIORPI', pfpm=False),
    'COR':  dict(suf='-COR', variante='Corredurías Públicas', sujeto='la Correduría',
                 frac='Fracción XII, Apartado B, del art. 17 LFPIORPI', pfpm=False),
}

# (número, título completo, título corto para el índice)
CAPITULOS = [
    (1,  'Introducción y Marco Normativo', 'Introducción y marco normativo'),
    (2,  'Objeto y Alcance', 'Objeto y alcance'),
    (3,  'Organización y Responsabilidades (art. 37 Bis, fr. X, RCG)', 'Organización y responsabilidades'),
    (4,  'Identificación y Conocimiento del Cliente (art. 37 Bis, fr. I, RCG)', 'Identificación y conocimiento del cliente'),
    (5,  'Enfoque Basado en Riesgo y Debida Diligencia (art. 37 Bis, fr. II y III, RCG)', 'Enfoque Basado en Riesgo y debida diligencia'),
    (6,  "Personas Políticamente Expuestas (PEP's) (art. 37 Bis, fr. IV, RCG)", "Personas Políticamente Expuestas"),
    (7,  'Beneficiario Controlador (art. 23 Quinquies RCG)', 'Beneficiario Controlador'),
    (8,  'Operaciones Inusuales (art. 18, fr. VI, de la Ley)', 'Operaciones inusuales'),
    (9,  'Restricción de Pagos en Efectivo (art. 32 de la Ley)', 'Restricción de pagos en efectivo'),
    (10, 'Presentación de Avisos e Informes (art. 37 Bis, fr. VI, RCG)', 'Avisos e Informes'),
    (11, 'Resguardo Documental (art. 37 Bis, fr. VII, RCG)', 'Resguardo documental'),
    (12, 'Capacitación (art. 37 Bis, fr. XI, RCG)', 'Capacitación'),
    (13, 'Control Interno, Supervisión y Auditoría (art. 37 Bis, fr. XII, RCG)', 'Control interno, supervisión y auditoría'),
    (14, 'Perfil Transaccional y Sistema de Alertas (art. 37 Bis, fr. V, RCG)', 'Perfil transaccional y alertas'),
    (15, 'Seguimiento y Acumulación de Actos u Operaciones (art. 37 Bis, fr. VIII, RCG)', 'Seguimiento y acumulación'),
    (16, 'Identificación de Personas en Listas de Autoridades (art. 37 Bis, fr. IX, RCG)', 'Listas de autoridades'),
    (17, 'Confidencialidad y Protección de Datos (art. 37 Bis, fr. XIII, RCG)', 'Confidencialidad y datos personales'),
    (18, 'Criterios para Aceptar, Limitar o Terminar la Relación (art. 37, tercer párrafo, RCG)', 'Aceptación, limitación y terminación'),
    (19, 'Mecanismos Automatizados (art. 41 RCG)', 'Mecanismos automatizados'),
    (20, 'Selección del Personal (art. 39 Bis 2 RCG)', 'Selección del personal'),
    (21, None, None),  # capítulo propio de cada sector (ver CAP21)
    (22, 'Actualización del Manual (art. 37 Bis, fr. XIV, RCG)', 'Actualización del Manual'),
    (23, 'Glosario', 'Glosario'),
]
CAP21 = {
    'PROF': ('Alcance de la Sujeción y Secreto Profesional (art. 17, fr. XI, de la Ley)', 'Alcance y secreto profesional'),
    'INM':  ('Alcance de la Sujeción por Actividad (art. 17, fr. V, V Bis y XV, de la Ley)', 'Alcance por actividad'),
    'NOT':  ('Disposiciones Propias de la Fe Pública Notarial (arts. 13 y 13 Bis RCG)', 'Disposiciones propias de la fe pública notarial'),
    'COR':  ('Disposiciones Propias de la Fe Pública del Corredor (arts. 13 y 13 Bis RCG)', 'Disposiciones propias de la fe pública del corredor'),
}
ANEXOS = [
    ('A', 'Correspondencia con el Contenido Mínimo del Manual (arts. 37 y 37 Bis RCG)', 'Correspondencia con el art. 37 Bis'),
    ('B', 'Cuestionario de Identificación Reforzada (arts. 23 Ter 3 y 23 Ter 4 RCG)', 'Cuestionario de Identificación Reforzada'),
    ('C', 'Descripción Operativa', 'Descripción Operativa'),
    ('D', 'Control de Versiones y Acuse de Difusión', 'Control de versiones y difusión'),
]

def titulo(kit, n):
    for num, t, c in CAPITULOS:
        if num == n:
            return (CAP21[kit][0] if n == 21 else t)
def corto(kit, n):
    for num, t, c in CAPITULOS:
        if num == n:
            return (CAP21[kit][1] if n == 21 else c)

def indice(kit):
    caps = ' │ '.join(f'{n}. {corto(kit, n)}' for n, _, _ in CAPITULOS)
    anx = ' │ '.join(f'{l}. {c}' for l, _, c in ANEXOS)
    return f'Índice: {caps} │ Anexos: {anx}'

def anexo_a(kit):
    """Filas del Anexo A (correspondencia con el contenido mínimo), iguales en los cuatro kits
    salvo el sufijo de los documentos y los supuestos propios del sector."""
    s = KITS[kit]['suf']
    d = lambda *xs: ', '.join(f'{x}{s}' for x in xs)
    d2 = {'PROF': 'D2-PROF-PF / D2-PROF-PM', 'INM': 'D2-PF / D2', 'NOT': 'D2-NOT', 'COR': 'D2-COR'}[kit]
    fila_x = {
        'PROF': 'Funciones y responsabilidades del Representante Encargado de Cumplimiento o, en persona física, del Profesional Titular',
        'INM': 'Funciones y responsabilidades del Representante Encargado de Cumplimiento o, en persona física, del titular',
        'NOT': 'Funciones y responsabilidades, ejercidas personalmente por el Notario Titular (art. 20, último párrafo, de la Ley)',
        'COR': 'Funciones y responsabilidades, ejercidas personalmente por el Corredor Público Titular (art. 20, último párrafo, de la Ley)',
    }[kit]
    fila_viii = {
        'PROF': 'Seguimiento de actos u operaciones (la fracción XI no acumula montos, art. 19 RCG)',
        'INM': 'Seguimiento y acumulación de actos u operaciones',
        'NOT': 'Seguimiento y acumulación de actos (incisos a), c) y d) del Apartado A, art. 19 RCG)',
        'COR': 'Seguimiento de actos (sin acumulación en el Apartado B, art. 19 RCG)',
    }[kit]
    kyc = d('D7') + (', D14' if kit == 'INM' else '')
    i_caps = 'Caps. 4, 7 y 21' if kit in ('NOT', 'COR') else 'Caps. 4 y 7'
    filas = [
        ('Fr.', 'Contenido mínimo', 'Capítulo de este Manual', 'Documento del Programa'),
        ('I', 'Identificación y conocimiento de Clientes o Usuarias, incluido el Beneficiario Controlador', i_caps, kyc),
        ('II', 'Mecanismos de clasificación de Riesgo', 'Cap. 5 (5.1)', d('D6')),
        ('III', 'Medidas de debida diligencia según el Grado de Riesgo', 'Cap. 5 (5.2 a 5.4); Anexo B', d('D6', 'D7')),
        ('IV', 'Identificación y seguimiento reforzado de PEP', 'Cap. 6', d('D8')),
        ('V', 'Detección de actos u operaciones que se aparten del Perfil Transaccional', 'Cap. 14', d('D9')),
        ('VI', 'Presentación de Avisos e Informes', 'Caps. 8 y 10', d('D9', 'D12')),
        ('VII', 'Conservación de información y documentación', 'Cap. 11', d('D10')),
        ('VIII', fila_viii, 'Cap. 15', d('D12')),
        ('IX', 'Identificación de personas en listas de autoridades', 'Cap. 16', d('D7')),
        ('X', fila_x, 'Cap. 3', d2),
        ('XI', 'Programas de capacitación', 'Cap. 12', d('D11', 'D11B', 'D11C')),
        ('XII', 'Control interno, supervisión y auditoría', 'Cap. 13', d('D13')),
        ('XIII', 'Confidencialidad de la información' + (' y secreto profesional' if kit == 'PROF' else ''),
         ('Caps. 17 y 21' if kit == 'PROF' else 'Cap. 17'), d('D10')),
        ('XIV', 'Actualización del Manual', 'Cap. 22; Anexo D', '—'),
        ('Art. 37, 3er párr.', 'Apertura, limitación y terminación de relaciones comerciales', 'Cap. 18', d('D6')),
        ('Art. 37 Bis 1', 'Grupo Empresarial',
         'Cap. 3 (3.3)' if kit == 'INM' else 'Cap. 3 (3.2)' if kit == 'PROF' else 'No aplica (persona física)', '—'),
        ('Art. 37 Bis 2', 'Supuestos que no se realizan (exención)', 'Cap. 2 (2.1)', '—'),
        ('Art. 37 Bis 3', 'Modificaciones ordenadas por el SAT', 'Cap. 22 (22.1)', '—'),
        ('Art. 10 Septies', 'Diseño de la metodología de evaluación de Riesgos', 'Cap. 5 (5.1)', d('D6')),
        ('Art. 17', 'Criterios de Riesgo bajo', 'Cap. 5 (5.2)', d('D6')),
        ('Art. 18', 'Actos u operaciones por medios electrónicos', 'Cap. 2 (2.3)', '—'),
        ('Art. 39 Bis 2', 'Selección de personal', 'Cap. 20', d('D11D')),
    ]
    if kit in ('NOT', 'COR'):
        filas.append(('Arts. 13 y 13 Bis', 'Solicitante material; fe pública al Poder Judicial o a autoridades administrativas',
                      'Cap. 21', d('D7')))
    return filas
