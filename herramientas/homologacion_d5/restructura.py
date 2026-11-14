"""Reorganiza plantillas ya homologadas en formato: secciones, encabezados, inserciones y correcciones de texto.

Uso típico:
    R = Doc(path, kit)
    R.renombrar(3, 'Riesgo …')            # título de la sección 3 (1-based)
    R.insertar_despues(5, 'Recomendación', lambda ctx: p(ctx, '…'))
    R.texto(lambda t: t.replace(…))       # por cada w:t
    R.guardar(dest)                        # renumera "Sección I, II…" y agrega paginación
"""
import os, sys, re, copy
sys.path.insert(0, os.path.dirname(__file__))
from build_d5 import Ctx, escribir_par, Par
from docxlib import q, numero_de_pagina, leer_par, espaciar_tras_tablas
from docx import Document

ROM = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII', 'XIII', 'XIV', 'XV', 'XVI']

def _estilo(el):
    ppr = el.find(q('w:pPr'))
    if ppr is None or ppr.find(q('w:pStyle')) is None:
        return ''
    return ppr.find(q('w:pStyle')).get(q('w:val'))

def _txt(el):
    return ''.join(t.text or '' for t in el.iter(q('w:t')))

class Doc:
    def __init__(self, path, kit):
        self.doc = Document(path)
        self.kit = kit
        self.body = self.doc.element.body

    def caps(self):
        return [el for el in self.body if el.tag == q('w:p') and _estilo(el) == 'CertusCapitulo' and re.match(r'Sección [IVX]+:', _txt(el))]

    def titulo(self, n):
        return re.sub(r'^Sección [IVX]+:\s*', '', _txt(self.caps()[n - 1]))

    def renombrar(self, n, titulo):
        el = self.caps()[n - 1]
        nuevo = escribir_par(Par.from_text('Sección X: %s' % titulo), 'CertusCapitulo')
        el.addprevious(nuevo); self.body.remove(el)

    def _fin_seccion(self, n):
        caps = self.caps()
        if n < len(caps):
            return caps[n]            # insertar antes del siguiente encabezado
        # última sección: antes del pie (CertusPie) o de sectPr
        for el in self.body:
            if el.tag == q('w:p') and _estilo(el) == 'CertusPie':
                return el
        return self.body.find(q('w:sectPr'))

    def _ctx(self, fn):
        ctx = Ctx(self.kit); fn(ctx); espaciar_tras_tablas(ctx.body); return ctx.body

    def insertar_despues(self, n, titulo, fn):
        """Nueva sección después de la sección n (0 = antes de la primera)."""
        ancla = self.caps()[0] if n == 0 else self._fin_seccion(n)
        els = [escribir_par(Par.from_text('Sección X: %s' % titulo), 'CertusCapitulo')] + self._ctx(fn)
        for el in els:
            ancla.addprevious(el)

    def agregar_a(self, n, fn):
        """Agrega contenido al final de la sección n."""
        ancla = self._fin_seccion(n)
        for el in self._ctx(fn):
            ancla.addprevious(el)

    def reemplazar_cuerpo(self, n, fn):
        """Sustituye todo el contenido de la sección n (conserva el encabezado)."""
        cap = self.caps()[n - 1]; fin = self._fin_seccion(n)
        el = cap.getnext()
        while el is not None and el is not fin:
            sig = el.getnext(); self.body.remove(el); el = sig
        for e in self._ctx(fn):
            fin.addprevious(e)

    def eliminar_parrafos(self, patron):
        n = 0
        for el in list(self.body):
            if el.tag == q('w:p') and re.search(patron, _txt(el)):
                self.body.remove(el); n += 1
        return n

    def texto(self, fn):
        """Aplica fn a cada w:t del documento (cuerpo, tablas). Devuelve número de cambios."""
        n = 0
        for t in self.body.iter(q('w:t')):
            if t.text:
                nuevo = fn(t.text)
                if nuevo != t.text:
                    t.text = nuevo; n += 1
                    if nuevo != nuevo.strip():
                        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        return n

    def parrafos(self, fn, estilos=None):
        """Reescribe párrafos completos con fn(Par)->Par|None (None = sin cambio), conservando el estilo."""
        n = 0
        for el in list(self.body.iter(q('w:p'))):
            est = _estilo(el)
            if estilos and est not in estilos:
                continue
            par = leer_par(el)
            res = fn(par)
            if res is None:
                continue
            nuevo = escribir_par(res, est or 'CertusTexto')
            ppr_old = el.find(q('w:pPr'))
            if ppr_old is not None:
                ppr_new = nuevo.find(q('w:pPr'))
                if ppr_new is not None:
                    nuevo.remove(ppr_new)
                nuevo.insert(0, copy.deepcopy(ppr_old))
            el.addprevious(nuevo); el.getparent().remove(el); n += 1
        return n

    def renumerar(self):
        for i, el in enumerate(self.caps()):
            for t in el.iter(q('w:t')):
                if t.text and re.match(r'Sección [IVX]+:', t.text):
                    t.text = re.sub(r'^Sección [IVX]+:', 'Sección %s:' % ROM[i], t.text); break

    def guardar(self, dest):
        self.renumerar()
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        self.doc.save(dest)

def _degradar(self, n, titulo):
    """Convierte el encabezado de la sección n en subsección (CertusSeccion) con nuevo título."""
    el = self.caps()[n - 1]
    nuevo = escribir_par(Par.from_text(titulo), 'CertusSeccion')
    el.addprevious(nuevo); self.body.remove(el)
Doc.degradar = _degradar
