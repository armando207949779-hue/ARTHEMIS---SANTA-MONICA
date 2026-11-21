"""Generador editorial de plantillas de capacitación para OTEC Santa Mónica.

Uso:
    pip install streamlit python-docx reportlab pillow requests
    streamlit run plantilla_otec_profesional.py

La salida Word es totalmente editable. El PDF reproduce la misma identidad visual.
No se exige conexión a internet: el logo puede cargarse localmente y la tipografía
PDF utiliza fuentes del sistema cuando están disponibles.
"""

from __future__ import annotations

import io
import re
from datetime import date
from html import escape
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from PIL import Image
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph


# -------------------------------- Configuración --------------------------------
LOGO_LOCAL = "OTEC_Santa_Monica_SII.jpg"
LOGO_URL = (
    "https://raw.githubusercontent.com/armando207949779-hue/"
    "ARTHEMIS---SANTA-MONICA/main/OTEC_Santa_Monica_SII.jpg"
)

PALETAS = {
    "Azul marino y turquesa": ("#183B56", "#169A92"),
    "Petróleo y cobre": ("#174B50", "#BD8552"),
    "Grafito y dorado": ("#273245", "#AB8650"),
    "Burdeo y terracota": ("#592F3C", "#BC785B"),
}

TINTA = "#263547"
GRIS = "#627287"
BORDE = "#D9E2EA"
FONDO = "#F4F7F9"
BLANCO = "#FFFFFF"


def rgb(h: str) -> tuple[int, int, int]:
    h = h.lstrip("#")
    if not re.fullmatch(r"[0-9a-fA-F]{6}", h):
        raise ValueError(f"Color inválido: {h}")
    return tuple(int(h[i : i + 2], 16) for i in (0, 2, 4))


def rlcolor(h: str):
    return colors.HexColor(h)


def texto(valor: object) -> str:
    """Normaliza controles sin eliminar tildes ni caracteres españoles."""
    return " ".join(str(valor or "").split()).strip()


def corto(valor: object, limite: int) -> str:
    valor = texto(valor)
    if len(valor) <= limite:
        return valor
    return valor[: max(0, limite - 1)].rstrip() + "…"


def validar(datos: dict):
    obligatorios = ("otec_nombre", "nombre_curso", "tipo_documento")
    for clave in obligatorios:
        if not texto(datos.get(clave)):
            raise ValueError(f"Falta completar el campo: {clave.replace('_', ' ')}")
    if not 1 <= int(datos["hojas_blancas"]) <= 30:
        raise ValueError("El número de hojas interiores debe estar entre 1 y 30.")
    if len(texto(datos["nombre_curso"])) > 190:
        raise ValueError("Acorta el título del curso a 190 caracteres como máximo.")
    if len(texto(datos["titulo_guia"])) > 170:
        raise ValueError("Acorta el subtítulo a 170 caracteres como máximo.")
    for campo in ("color_primario", "color_acento"):
        rgb(datos[campo])


def logo_predeterminado() -> bytes | None:
    archivo = Path(LOGO_LOCAL)
    if archivo.is_file():
        try:
            return archivo.read_bytes()
        except OSError:
            pass
    try:
        import requests

        respuesta = requests.get(LOGO_URL, timeout=6)
        respuesta.raise_for_status()
        if len(respuesta.content) <= 5_000_000:
            Image.open(io.BytesIO(respuesta.content)).verify()
            return respuesta.content
    except Exception:
        pass
    return None


def imagen_dimensiones(raw: bytes) -> tuple[int, int]:
    with Image.open(io.BytesIO(raw)) as imagen:
        return imagen.size


# ----------------------------- Diseño compartido ------------------------------
def ficha(datos: dict) -> list[tuple[str, str]]:
    vacio = "________________" if datos.get("linea_vacia") else "Por completar"
    return [
        ("RELATOR / DOCENTE", texto(datos.get("relator")) or vacio),
        ("INSTITUCIÓN / CLIENTE", texto(datos.get("cliente")) or vacio),
        ("MODALIDAD", texto(datos.get("modalidad")) or vacio),
        ("DURACIÓN", texto(datos.get("horas")) or vacio),
        ("EDICIÓN", texto(datos.get("codigo")) or vacio),
        ("FECHA", texto(datos.get("fecha")) or vacio),
    ]


def tamano_papel(datos: dict):
    return A4 if datos.get("papel") == "A4" else letter


def titulo_tamano(titulo: str) -> float:
    n = len(texto(titulo))
    return 29 if n < 54 else 26 if n < 85 else 22 if n < 130 else 19


# ------------------------------ Utilidades Word -------------------------------
def color_word(h: str):
    return RGBColor(*rgb(h))


def fondo_celda(celda, hex_color: str):
    tcpr = celda._tc.get_or_add_tcPr()
    anterior = tcpr.find(qn("w:shd"))
    if anterior is not None:
        tcpr.remove(anterior)
    nodo = OxmlElement("w:shd")
    nodo.set(qn("w:fill"), hex_color.lstrip("#"))
    tcpr.append(nodo)


def margenes_celda(celda, top=100, left=160, bottom=100, right=160):
    tcpr = celda._tc.get_or_add_tcPr()
    anterior = tcpr.find(qn("w:tcMar"))
    if anterior is not None:
        tcpr.remove(anterior)
    tc_mar = OxmlElement("w:tcMar")
    for lado, valor in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        nodo = OxmlElement(f"w:{lado}")
        nodo.set(qn("w:w"), str(valor))
        nodo.set(qn("w:type"), "dxa")
        tc_mar.append(nodo)
    tcpr.append(tc_mar)


def bordes_tabla(tabla, color=None):
    tblpr = tabla._tbl.tblPr
    anterior = tblpr.find(qn("w:tblBorders"))
    if anterior is not None:
        tblpr.remove(anterior)
    bordes = OxmlElement("w:tblBorders")
    for borde in ("top", "left", "bottom", "right", "insideH", "insideV"):
        elemento = OxmlElement(f"w:{borde}")
        if color:
            elemento.set(qn("w:val"), "single")
            elemento.set(qn("w:sz"), "4")
            elemento.set(qn("w:color"), color.lstrip("#"))
        else:
            elemento.set(qn("w:val"), "nil")
        bordes.append(elemento)
    tblpr.append(bordes)


def borde_parrafo(p, color: str, lado="bottom", grosor="8"):
    ppr = p._p.get_or_add_pPr()
    antiguo = ppr.find(qn("w:pBdr"))
    if antiguo is not None:
        ppr.remove(antiguo)
    pbdr = OxmlElement("w:pBdr")
    el = OxmlElement(f"w:{lado}")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), grosor)
    el.set(qn("w:space"), "6")
    el.set(qn("w:color"), color.lstrip("#"))
    pbdr.append(el)
    ppr.append(pbdr)


def campo_word(p, nombre: str, color: str):
    run = p.add_run()
    run.font.name = "Arial"
    run.font.size = Pt(8)
    run.font.bold = True
    run.font.color.rgb = color_word(color)
    inicio = OxmlElement("w:fldChar")
    inicio.set(qn("w:fldCharType"), "begin")
    codigo = OxmlElement("w:instrText")
    codigo.set(qn("xml:space"), "preserve")
    codigo.text = f" {nombre} "
    separador = OxmlElement("w:fldChar")
    separador.set(qn("w:fldCharType"), "separate")
    cache = OxmlElement("w:t")
    cache.text = "1"
    fin = OxmlElement("w:fldChar")
    fin.set(qn("w:fldCharType"), "end")
    for elemento in (inicio, codigo, separador, cache, fin):
        run._r.append(elemento)


def texto_word(p, contenido, size=10, color=TINTA, bold=False):
    # Conserva el espacio final necesario cuando el siguiente run es un campo PAGE.
    origen = str(contenido if contenido is not None else "")
    valor = texto(origen)
    if origen.endswith(" ") and valor:
        valor += " "
    run = p.add_run(valor)
    run.font.name = "Arial"
    run.font.size = Pt(size)
    run.font.color.rgb = color_word(color)
    run.bold = bold
    return run


def logo_word(p, raw: bytes | None, max_ancho=1.25, max_alto=0.62, fallback="SANTA MÓNICA"):
    if raw:
        try:
            ancho, alto = imagen_dimensiones(raw)
            factor = min(max_ancho / ancho, max_alto / alto)
            p.add_run().add_picture(io.BytesIO(raw), width=Inches(ancho * factor))
            return
        except Exception:
            pass
    texto_word(p, fallback, size=10, color=TINTA, bold=True)


def preparar_estilos(doc: Document, datos: dict):
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(11)
    normal.font.color.rgb = color_word(TINTA)
    normal.paragraph_format.space_after = Pt(8)
    normal.paragraph_format.line_spacing = 1.18

    for nombre, tam, antes, despues in (
        ("Heading 1", 16, 19, 8),
        ("Heading 2", 13, 14, 6),
        ("Heading 3", 11, 11, 4),
    ):
        estilo = doc.styles[nombre]
        estilo.font.name = "Arial"
        estilo.font.size = Pt(tam)
        estilo.font.bold = True
        estilo.font.color.rgb = color_word(datos["color_primario"])
        estilo.paragraph_format.space_before = Pt(antes)
        estilo.paragraph_format.space_after = Pt(despues)
        estilo.paragraph_format.keep_with_next = True


def generar_word(datos: dict, logo_bytes: bytes | None = None) -> bytes:
    """Crea portada de una página e interiores editables numerados desde 1."""
    validar(datos)
    doc = Document()
    preparar_estilos(doc, datos)
    principal, acento = datos["color_primario"], datos["color_acento"]

    # La portada usa su propia sección, sin encabezado ni pie repetidos.
    portada = doc.sections[0]
    portada.page_width, portada.page_height = (Inches(8.27), Inches(11.69)) if datos.get("papel") == "A4" else (Inches(8.5), Inches(11))
    portada.top_margin = Inches(0.64)
    portada.bottom_margin = Inches(0.56)
    portada.left_margin = Inches(0.74)
    portada.right_margin = Inches(0.74)
    portada.different_first_page_header_footer = True
    ancho = portada.page_width.inches - portada.left_margin.inches - portada.right_margin.inches

    # Marca y datos de edición.
    cab = doc.add_table(rows=1, cols=2)
    cab.autofit = False
    cab.columns[0].width = Inches(ancho * 0.58)
    cab.columns[1].width = Inches(ancho * 0.42)
    bordes_tabla(cab)
    izq, der = cab.rows[0].cells
    margenes_celda(izq, 40, 0, 40, 40)
    margenes_celda(der, 40, 40, 40, 0)
    logo_word(izq.paragraphs[0], logo_bytes, max_ancho=1.75, max_alto=0.78,
              fallback="OTEC SANTA MÓNICA")
    p = der.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    texto_word(p, corto(datos.get("codigo"), 35), 9, principal, True)
    p2 = der.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.paragraph_format.space_after = Pt(0)
    texto_word(p2, datos.get("fecha"), 9, GRIS)

    linea = doc.add_paragraph()
    linea.paragraph_format.space_before = Pt(8)
    linea.paragraph_format.space_after = Pt(0)
    borde_parrafo(linea, acento, grosor="16")

    espacio = doc.add_paragraph()
    espacio.paragraph_format.space_after = Pt(22)

    tipo = doc.add_paragraph()
    tipo.paragraph_format.space_after = Pt(12)
    texto_word(tipo, texto(datos["tipo_documento"]).upper(), 10, acento, True)

    titulo = doc.add_paragraph()
    titulo.paragraph_format.space_after = Pt(12)
    titulo.paragraph_format.line_spacing = 1.06
    titulo.paragraph_format.keep_with_next = True
    texto_word(titulo, datos["nombre_curso"], titulo_tamano(datos["nombre_curso"]), principal, True)

    sub = doc.add_paragraph()
    sub.paragraph_format.space_after = Pt(8)
    sub.paragraph_format.line_spacing = 1.16
    texto_word(sub, datos.get("titulo_guia"), 14, GRIS)

    # Ficha con etiquetas consistentes, alturas automáticas y sin recortes.
    espacio2 = doc.add_paragraph()
    espacio2.paragraph_format.space_before = Pt(8)
    espacio2.paragraph_format.space_after = Pt(9)
    etiqueta = doc.add_paragraph()
    etiqueta.paragraph_format.space_after = Pt(7)
    texto_word(etiqueta, "INFORMACIÓN DEL DOCUMENTO", 9, acento, True)

    meta = doc.add_table(rows=3, cols=2)
    meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta.autofit = False
    meta.columns[0].width = Inches(ancho / 2)
    meta.columns[1].width = Inches(ancho / 2)
    bordes_tabla(meta, BORDE)
    for i, (etiqueta, valor) in enumerate(ficha(datos)):
        celda = meta.cell(i // 2, i % 2)
        fondo_celda(celda, FONDO if i // 2 % 2 == 0 else BLANCO)
        margenes_celda(celda, top=110, left=170, bottom=100, right=130)
        celda.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = celda.paragraphs[0]
        p.paragraph_format.space_after = Pt(3)
        texto_word(p, etiqueta, 7.7, acento, True)
        p2 = celda.add_paragraph()
        p2.paragraph_format.space_after = Pt(0)
        texto_word(p2, valor, 10, TINTA)

    cierre = doc.add_paragraph()
    cierre.paragraph_format.space_before = Pt(18)
    cierre.paragraph_format.space_after = Pt(0)
    texto_word(cierre, datos["otec_nombre"], 9, principal, True)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    texto_word(p, datos.get("pie_pagina"), 8, GRIS)

    # Sección interior: no hereda encabezados ni pies de la portada.
    interior = doc.add_section(WD_SECTION.NEW_PAGE)
    interior.page_width, interior.page_height = portada.page_width, portada.page_height
    interior.top_margin = Inches(1.12)
    interior.bottom_margin = Inches(0.93)
    interior.left_margin = Inches(0.83)
    interior.right_margin = Inches(0.83)
    interior.header_distance = Inches(0.35)
    interior.footer_distance = Inches(0.39)
    interior.different_first_page_header_footer = False
    interior.header.is_linked_to_previous = False
    interior.footer.is_linked_to_previous = False
    ancho_int = interior.page_width.inches - interior.left_margin.inches - interior.right_margin.inches

    # Reinicia PAGE en 1 dentro de la sección. Sin un total fijo, Word sigue
    # numerando correctamente cuando el relator añade contenido.
    sectpr = interior._sectPr
    pgnum = sectpr.find(qn("w:pgNumType"))
    if pgnum is None:
        pgnum = OxmlElement("w:pgNumType")
        sectpr.append(pgnum)
    pgnum.set(qn("w:start"), "1")

    encabezado = interior.header
    tab = encabezado.add_table(rows=1, cols=2, width=Inches(ancho_int))
    tab.autofit = False
    tab.columns[0].width = Inches(ancho_int * 0.26)
    tab.columns[1].width = Inches(ancho_int * 0.74)
    bordes_tabla(tab)
    c_logo, c_info = tab.rows[0].cells
    margenes_celda(c_logo, 0, 0, 0, 0)
    margenes_celda(c_info, 0, 40, 0, 0)
    logo_word(c_logo.paragraphs[0], logo_bytes, max_ancho=1.2, max_alto=0.39,
              fallback="OTEC SANTA MÓNICA")
    p = c_info.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_after = Pt(1)
    texto_word(p, corto(datos["nombre_curso"], 64), 8.3, principal, True)
    p = c_info.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_after = Pt(0)
    texto_word(p, corto(datos.get("titulo_guia"), 66), 7.4, GRIS)
    regla = encabezado.add_paragraph()
    regla.paragraph_format.space_before = Pt(0)
    regla.paragraph_format.space_after = Pt(0)
    borde_parrafo(regla, acento, grosor="7")

    pie = interior.footer
    rule = pie.paragraphs[0]
    rule.paragraph_format.space_after = Pt(5)
    borde_parrafo(rule, BORDE, lado="top", grosor="6")
    tf = pie.add_table(rows=1, cols=2, width=Inches(ancho_int))
    tf.autofit = False
    tf.columns[0].width = Inches(ancho_int * 0.72)
    tf.columns[1].width = Inches(ancho_int * 0.28)
    bordes_tabla(tf)
    p = tf.cell(0, 0).paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    texto_word(p, corto(datos["otec_nombre"], 49), 7.3, principal, True)
    p2 = tf.cell(0, 0).add_paragraph()
    p2.paragraph_format.space_after = Pt(0)
    texto_word(p2, corto(datos.get("pie_pagina"), 82), 7, GRIS)
    num = tf.cell(0, 1).paragraphs[0]
    num.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    num.paragraph_format.space_after = Pt(0)
    texto_word(num, "PÁG.  ", 8, GRIS)
    campo_word(num, "PAGE", principal)

    # Las páginas limpias ya llevan su formato; no contienen párrafos de relleno.
    if datos.get("estructura_inicial"):
        h = doc.add_paragraph("Objetivo de aprendizaje", style="Heading 1")
        h.paragraph_format.keep_with_next = True
        doc.add_paragraph("[Redacte aquí el objetivo de aprendizaje del módulo.]")
        doc.add_paragraph("Contenidos", style="Heading 1")
        doc.add_paragraph("[Desarrolle los contenidos del curso.]")
        doc.add_paragraph("Actividad práctica", style="Heading 1")
        doc.add_paragraph("[Indique instrucciones, materiales y tiempo disponible.]")
        doc.add_paragraph("Evaluación y cierre", style="Heading 1")
        doc.add_paragraph("[Describa evidencias y criterios de evaluación.]")
    else:
        p = doc.add_paragraph(" ")
        p.paragraph_format.space_after = Pt(0)

    for _ in range(int(datos["hojas_blancas"]) - 1):
        doc.add_page_break()
        p = doc.add_paragraph(" ")
        p.paragraph_format.space_after = Pt(0)

    # Word recalcula los campos al abrir el archivo.
    settings = doc.settings.element
    update = settings.find(qn("w:updateFields"))
    if update is None:
        update = OxmlElement("w:updateFields")
        settings.append(update)
    update.set(qn("w:val"), "true")

    salida = io.BytesIO()
    doc.save(salida)
    return salida.getvalue()


# ------------------------------ Utilidades PDF --------------------------------
def fuentes_pdf():
    """Intenta incrustar una fuente Unicode; evita perder tildes y símbolos."""
    posibles = [
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
         "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf"),
        ("/Library/Fonts/Arial Unicode.ttf", "/Library/Fonts/Arial Bold.ttf"),
        ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
    ]
    for regular, negrita in posibles:
        if Path(regular).exists() and Path(negrita).exists():
            try:
                if "OTEC-Regular" not in pdfmetrics.getRegisteredFontNames():
                    pdfmetrics.registerFont(TTFont("OTEC-Regular", regular))
                    pdfmetrics.registerFont(TTFont("OTEC-Bold", negrita))
                return "OTEC-Regular", "OTEC-Bold"
            except Exception:
                continue
    return "Helvetica", "Helvetica-Bold"


def pdf_parrafo(c, contenido: str, x: float, y_top: float, ancho: float, alto_max: float,
                fuente: str, color: str, size: float, min_size: float = 7,
                leading_factor: float = 1.25, align=TA_LEFT) -> float:
    """Dibuja texto de arriba abajo, con salto de línea y tamaño adaptable."""
    contenido = escape(texto(contenido)) or " "
    actual = size
    while actual >= min_size - 0.01:
        estilo = ParagraphStyle(
            "OTEC", fontName=fuente, fontSize=actual,
            leading=actual * leading_factor, textColor=rlcolor(color), alignment=align,
            spaceBefore=0, spaceAfter=0, wordWrap="CJK",
        )
        p = Paragraph(contenido, estilo)
        _, alto = p.wrap(ancho, alto_max + 500)
        if alto <= alto_max:
            p.drawOn(c, x, y_top - alto)
            return alto
        actual -= 0.5
    raise ValueError(f"Texto demasiado largo para el diseño: {texto(contenido)[:48]}")


def pdf_logo(c, raw: bytes | None, x, y_top, max_w, max_h,
             principal: str, bold_font: str):
    if raw:
        try:
            w, h = imagen_dimensiones(raw)
            escala = min(max_w / w, max_h / h)
            dw, dh = w * escala, h * escala
            c.drawImage(ImageReader(io.BytesIO(raw)), x, y_top - dh,
                        width=dw, height=dh, mask="auto")
            return
        except Exception:
            pass
    pdf_parrafo(c, "OTEC SANTA MÓNICA", x, y_top, max_w, max_h,
                bold_font, principal, 11, 8)


def generar_pdf(datos: dict, logo_bytes: bytes | None = None) -> bytes:
    validar(datos)
    buffer = io.BytesIO()
    w, h = tamano_papel(datos)
    c = canvas.Canvas(buffer, pagesize=(w, h), pageCompression=1)
    c.setTitle(texto(datos["nombre_curso"]))
    c.setAuthor(texto(datos["otec_nombre"]))
    normal, bold = fuentes_pdf()
    principal, acento = datos["color_primario"], datos["color_acento"]
    x = 52
    ancho = w - 104

    # ------------------------------- Portada ---------------------------------
    pdf_logo(c, logo_bytes, x, h - 50, 146, 56, principal, bold)
    pdf_parrafo(c, corto(datos.get("codigo"), 38), w - 225, h - 53, 173, 19,
                bold, principal, 9, 7, align=TA_RIGHT)
    pdf_parrafo(c, datos.get("fecha"), w - 225, h - 72, 173, 16,
                normal, GRIS, 9, 7, align=TA_RIGHT)

    c.setStrokeColor(rlcolor(acento))
    c.setLineWidth(2.1)
    c.line(x, h - 124, w - x, h - 124)

    pdf_parrafo(c, texto(datos["tipo_documento"]).upper(), x, h - 187,
                ancho, 22, bold, acento, 11, 9)
    altura_tit = pdf_parrafo(c, datos["nombre_curso"], x, h - 221,
                            ancho, 136, bold, principal, titulo_tamano(datos["nombre_curso"]), 16,
                            leading_factor=1.13)
    pdf_parrafo(c, datos.get("titulo_guia"), x, h - 234 - altura_tit,
                ancho, 60, normal, GRIS, 15, 10, leading_factor=1.22)

    # Malla de información anclada abajo, incluso cuando varía el título.
    y_label = 273
    pdf_parrafo(c, "INFORMACIÓN DEL DOCUMENTO", x, y_label + 26,
                ancho, 15, bold, acento, 9.2, 8)
    alto = 46
    sep = 7
    col_w = (ancho - sep) / 2
    for i, (etiqueta, valor) in enumerate(ficha(datos)):
        fila, col = divmod(i, 2)
        bx = x + col * (col_w + sep)
        by = y_label - fila * (alto + 7)
        c.setFillColor(rlcolor(FONDO if fila % 2 == 0 else BLANCO))
        c.setStrokeColor(rlcolor(BORDE))
        c.setLineWidth(0.65)
        c.roundRect(bx, by - alto, col_w, alto, radius=3, stroke=1, fill=1)
        c.setFillColor(rlcolor(acento))
        c.rect(bx, by - alto, 3, alto, fill=1, stroke=0)
        pdf_parrafo(c, etiqueta, bx + 12, by - 9, col_w - 24, 13,
                    bold, acento, 7.7, 6.5)
        pdf_parrafo(c, valor, bx + 12, by - 22, col_w - 24, 23,
                    normal, TINTA, 9.8, 7.5)

    pdf_parrafo(c, datos["otec_nombre"], x, 91, ancho, 18,
                bold, principal, 9, 7)
    pdf_parrafo(c, datos.get("pie_pagina"), x, 76, ancho, 24,
                normal, GRIS, 8, 6.5)
    c.showPage()

    # ------------------------ Interiores sin portada -------------------------
    left = 58
    usable = w - left * 2
    for numero in range(1, int(datos["hojas_blancas"]) + 1):
        # Encabezado estable, sin invadir el área de escritura.
        pdf_logo(c, logo_bytes, left, h - 30, 110, 33, principal, bold)
        pdf_parrafo(c, corto(datos["nombre_curso"], 70), left + 125, h - 31,
                    usable - 125, 18, bold, principal, 8.6, 7, align=TA_RIGHT)
        pdf_parrafo(c, corto(datos.get("titulo_guia"), 66), left + 125, h - 49,
                    usable - 125, 18, normal, GRIS, 7.6, 6.5, align=TA_RIGHT)
        c.setStrokeColor(rlcolor(acento))
        c.setLineWidth(0.75)
        c.line(left, h - 79, w - left, h - 79)

        if numero == 1 and datos.get("estructura_inicial"):
            y = h - 129
            for cabecera, descripcion in (
                ("Objetivo de aprendizaje", "[Redacte aquí el objetivo de aprendizaje del módulo.]"),
                ("Contenidos", "[Desarrolle los contenidos del curso.]"),
                ("Actividad práctica", "[Indique instrucciones, materiales y tiempo disponible.]"),
                ("Evaluación y cierre", "[Describa evidencias y criterios de evaluación.]"),
            ):
                pdf_parrafo(c, cabecera, left, y, usable, 35, bold, principal, 14, 12)
                y -= 29
                pdf_parrafo(c, descripcion, left, y, usable, 40, normal, GRIS, 10, 8)
                y -= 74

        # Pie de página con numeración relativa a los interiores.
        c.setStrokeColor(rlcolor(BORDE))
        c.setLineWidth(0.6)
        c.line(left, 57, w - left, 57)
        pdf_parrafo(c, corto(datos["otec_nombre"], 62), left, 44,
                    usable * 0.68, 13, bold, principal, 7.7, 6.6)
        pdf_parrafo(c, corto(datos.get("pie_pagina"), 85), left, 31,
                    usable * 0.72, 18, normal, GRIS, 7, 6.1)
        pdf_parrafo(c, f"PÁG. {numero}",
                    left + usable * 0.72, 43, usable * 0.28, 20,
                    bold, principal, 8, 7, align=TA_RIGHT)
        c.showPage()

    c.save()
    return buffer.getvalue()


# ----------------------------- Interfaz Streamlit ------------------------------
def run():
    try:
        import streamlit as st
    except ImportError as exc:
        raise RuntimeError(
            "Falta Streamlit. Instala las dependencias con: "
            "pip install streamlit python-docx reportlab pillow requests"
        ) from exc

    st.markdown(
        """<style>
        .block-container {max-width: 1150px; padding-top: 2.1rem}
        h1, h2, h3 {letter-spacing: -.025em}
        div.stButton > button[kind=primary] {min-height: 3rem; font-weight: 650}
        </style>""",
        unsafe_allow_html=True,
    )
    st.title("Plantillas de capacitación")
    st.caption("OTEC Santa Mónica · Portada institucional, páginas interiores y pie corporativo")

    with st.sidebar:
        st.header("Diseño institucional")
        paleta = st.selectbox("Paleta de colores", list(PALETAS) + ["Personalizada"])
        if paleta == "Personalizada":
            primario = st.color_picker("Color principal", "#183B56")
            acento = st.color_picker("Color de acento", "#169A92")
        else:
            primario, acento = PALETAS[paleta]
            st.markdown(f"**Principal:** `{primario}`  **Acento:** `{acento}`")
        papel = st.selectbox("Formato del papel", ["Carta", "A4"])
        st.markdown("---")
        logo_subido = st.file_uploader("Logo de la institución", type=["png", "jpg", "jpeg"])
        if logo_subido:
            logo = logo_subido.getvalue()
        else:
            @st.cache_data(show_spinner=False, ttl=3600)
            def cargar():
                return logo_predeterminado()
            logo = cargar()
        if logo:
            st.image(logo, width=150)
        else:
            st.info("Sin logo: se mostrará el nombre institucional.")
        st.caption("El documento es editable en Word; el PDF queda listo para distribuir.")

    t1, t2 = st.tabs(["Datos del documento", "Formato y contenido"])
    with t1:
        col1, col2 = st.columns(2, gap="large")
        with col1:
            otec_nombre = st.text_input("Nombre de la OTEC", "Centro de Capacitación Santa Mónica Limitada")
            tipo = st.selectbox("Tipo de documento", [
                "Manual del participante", "Guía de capacitación y relatoría",
                "Guía de taller práctico", "Material de apoyo docente",
            ])
            curso = st.text_input("Nombre del curso", "Inteligencia artificial y análisis de datos aplicado")
            subtitulo = st.text_input("Módulo o subtítulo", "Módulo 1: Fundamentos y herramientas prácticas")
        with col2:
            relator = st.text_input("Relator o docente", "")
            cliente = st.text_input("Institución o cliente", "")
            modalidad = st.selectbox("Modalidad", ["Presencial", "Online sincrónico", "E-learning", "Mixta"])
            horas = st.text_input("Duración", "16 horas cronológicas")
            codigo = st.text_input("Código / versión", "GUIA-01 / v1.0")
            fecha_sel = st.date_input("Fecha de edición", date.today(), format="DD/MM/YYYY")

    with t2:
        pie = st.text_input(
            "Texto institucional del pie",
            "Material exclusivo de capacitación - Prohibida su reproducción sin autorización",
        )
        hojas = st.slider("Cantidad de hojas interiores", min_value=1, max_value=30, value=2)
        estructura = st.checkbox(
            "Incluir estructura inicial editable en la primera hoja",
            value=False,
            help="Agrega Objetivo, Contenidos, Actividad práctica y Evaluación en lugar de una primera hoja vacía.",
        )
        linea_vacia = st.checkbox("Mostrar líneas para datos no completados", value=False)
        st.info("Las hojas posteriores a la portada mantienen siempre el mismo encabezado y pie. ")

    datos = dict(
        otec_nombre=otec_nombre, tipo_documento=tipo, nombre_curso=curso,
        titulo_guia=subtitulo, relator=relator, cliente=cliente,
        modalidad=modalidad, horas=horas, codigo=codigo,
        fecha=fecha_sel.strftime("%d/%m/%Y"), pie_pagina=pie,
        hojas_blancas=hojas, linea_vacia=linea_vacia,
        color_primario=primario, color_acento=acento,
        papel=papel, estructura_inicial=estructura,
    )

    st.divider()
    if st.button("Generar Word y PDF", type="primary", use_container_width=True):
        try:
            with st.spinner("Diseñando portada y páginas interiores..."):
                word = generar_word(datos, logo)
                pdf = generar_pdf(datos, logo)
        except Exception as exc:
            st.error(f"No se pudo generar la plantilla: {exc}")
        else:
            base = re.sub(r"[^\w-]+", "_", texto(curso), flags=re.UNICODE).strip("_")[:75]
            st.session_state["plantilla_otec"] = (word, pdf, base or "Curso")
            st.success("Listo. Descarga la plantilla con el nuevo formato editorial.")

    if "plantilla_otec" in st.session_state:
        word, pdf, base = st.session_state["plantilla_otec"]
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "Descargar Word editable", word, f"Plantilla_{base}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with c2:
            st.download_button("Descargar PDF", pdf, f"Plantilla_{base}.pdf",
                               mime="application/pdf", use_container_width=True)
        try:
            import fitz  # Opcional: pip install pymupdf
            muestra = fitz.open(stream=pdf, filetype="pdf")
            pixmap = muestra[0].get_pixmap(matrix=fitz.Matrix(1.25, 1.25), alpha=False)
            st.markdown("#### Vista previa de la portada")
            st.image(pixmap.tobytes("png"), width=470)
        except ImportError:
            st.caption("Para ver la portada dentro de la app: pip install pymupdf")


if __name__ == "__main__":
    import streamlit as st

    st.set_page_config(page_title="Plantillas OTEC Santa Mónica", page_icon="📘", layout="wide")
    run()
