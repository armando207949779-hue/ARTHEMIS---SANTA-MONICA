import io
from datetime import date
import requests
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from fpdf import FPDF

# URL RAW del logo en GitHub (convertida desde /blob/main/ a raw.githubusercontent.com)
LOGO_URL_DEFAULT = (
    "https://raw.githubusercontent.com/armando207949779-hue/"
    "ARTHEMIS---SANTA-MONICA/main/OTEC_Santa_Monica_SII.jpg"
)

# Paleta corporativa
COLOR_NAVY_HEX = "1B365D"
COLOR_ACCENT_HEX = "2B6CB0"
COLOR_LIGHT_BG_HEX = "F4F6F9"
COLOR_BORDER_HEX = "CBD5E0"

COLOR_NAVY_RGB = (27, 54, 93)
COLOR_ACCENT_RGB = (43, 108, 176)
COLOR_DARK_TEXT_RGB = (45, 55, 72)
COLOR_MUTED_RGB = (113, 128, 150)
COLOR_LIGHT_BG_RGB = (244, 246, 249)


# ==============================================================================
# UTILIDADES GENERALES
# ==============================================================================
@st.cache_data(show_spinner=False)
def descargar_logo(url: str) -> bytes | None:
    try:
        resp = requests.get(url, timeout=10)
        if resp.status_code == 200 and resp.content:
            return resp.content
    except Exception:
        pass
    return None


def limpiar_latin1(texto: str) -> str:
    """Evita errores de codificación en fuentes estándar de FPDF2."""
    reemplazos = {
        "—": "-",
        "–": "-",
        "“": '"',
        "”": '"',
        "‘": "'",
        "’": "'",
        "•": "-",
        "…": "...",
    }
    for k, v in reemplazos.items():
        texto = texto.replace(k, v)
    return texto.encode("latin-1", "replace").decode("latin-1")


# ==============================================================================
# UTILIDADES XML PARA PYTHON-DOCX
# ==============================================================================
def set_cell_shading(cell, color_hex: str):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_table_borders(table, color_hex=COLOR_BORDER_HEX, sz="4"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color_hex}"/>'
        f'  <w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color_hex}"/>'
        f'  <w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color_hex}"/>'
        f'  <w:left w:val="none"/>'
        f'  <w:right w:val="none"/>'
        f'  <w:insideV w:val="none"/>'
        f"</w:tblBorders>"
    )
    tblPr.append(borders)


def remove_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="none"/><w:bottom w:val="none"/>'
        f'  <w:left w:val="none"/><w:right w:val="none"/>'
        f'  <w:insideH w:val="none"/><w:insideV w:val="none"/>'
        f"</w:tblBorders>"
    )
    tblPr.append(borders)


def add_paragraph_border(paragraph, side="bottom", color_hex=COLOR_ACCENT_HEX, sz="12"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    border = OxmlElement(f"w:{side}")
    border.set(qn("w:val"), "single")
    border.set(qn("w:sz"), sz)
    border.set(qn("w:space"), "4")
    border.set(qn("w:color"), color_hex)
    pBdr.append(border)
    pPr.append(pBdr)


def add_page_number_field(run):
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = "PAGE"
    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "separate")
    fldChar3 = OxmlElement("w:fldChar")
    fldChar3.set(qn("w:fldCharType"), "end")
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)


# ==============================================================================
# GENERADOR WORD (.DOCX)
# ==============================================================================
def generar_word(datos: dict, logo_bytes: bytes | None) -> bytes:
    doc = Document()

    # Configuración de página (Carta / Letter)
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(0.9)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    section.different_first_page_header_footer = True

    # Fuente base
    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Calibri"
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(*COLOR_DARK_TEXT_RGB)

    # ------------------------------------------------------------------
    # ENCABEZADO (Páginas 2 en adelante)
    # ------------------------------------------------------------------
    header = section.header
    htbl = header.add_table(rows=1, cols=2, width=Inches(6.5))
    htbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(htbl)

    c_logo, c_info = htbl.rows[0].cells
    c_logo.width = Inches(1.8)
    c_info.width = Inches(4.7)

    p_hlogo = c_logo.paragraphs[0]
    p_hlogo.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if logo_bytes:
        p_hlogo.add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(1.2))
    else:
        r = p_hlogo.add_run("OTEC SANTA MÓNICA")
        r.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(*COLOR_NAVY_RGB)

    p_hinfo = c_info.paragraphs[0]
    p_hinfo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r1 = p_hinfo.add_run(f"{datos['nombre_curso']}\n")
    r1.bold = True
    r1.font.size = Pt(9)
    r1.font.color.rgb = RGBColor(*COLOR_NAVY_RGB)
    r2 = p_hinfo.add_run(f"{datos['titulo_guia']}  |  {datos['otec_nombre']}")
    r2.font.size = Pt(8)
    r2.font.color.rgb = RGBColor(*COLOR_MUTED_RGB)

    # Línea divisoria bajo encabezado
    p_hline = header.add_paragraph()
    p_hline.paragraph_format.space_before = Pt(2)
    p_hline.paragraph_format.space_after = Pt(6)
    add_paragraph_border(p_hline, side="bottom", color_hex=COLOR_ACCENT_HEX, sz="8")

    # ------------------------------------------------------------------
    # PIE DE PÁGINA (Páginas 2 en adelante)
    # ------------------------------------------------------------------
    footer = section.footer
    p_fline = footer.paragraphs[0]
    p_fline.paragraph_format.space_before = Pt(0)
    p_fline.paragraph_format.space_after = Pt(4)
    add_paragraph_border(p_fline, side="top", color_hex=COLOR_BORDER_HEX, sz="6")

    ftbl = footer.add_table(rows=1, cols=2, width=Inches(6.5))
    ftbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(ftbl)

    fc1, fc2 = ftbl.rows[0].cells
    fc1.width = Inches(4.8)
    fc2.width = Inches(1.7)

    p_f1 = fc1.paragraphs[0]
    rf1 = p_f1.add_run(f"{datos['otec_nombre']}  |  {datos['pie_pagina']}")
    rf1.font.size = Pt(8)
    rf1.font.color.rgb = RGBColor(*COLOR_MUTED_RGB)

    p_f2 = fc2.paragraphs[0]
    p_f2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    rf2 = p_f2.add_run("Página ")
    rf2.font.size = Pt(8)
    rf2.font.color.rgb = RGBColor(*COLOR_MUTED_RGB)
    rf_num = p_f2.add_run()
    rf_num.font.size = Pt(8)
    rf_num.font.bold = True
    rf_num.font.color.rgb = RGBColor(*COLOR_NAVY_RGB)
    add_page_number_field(rf_num)

    # ------------------------------------------------------------------
    # PORTADA (Página 1)
    # ------------------------------------------------------------------
    # Franja superior decorativa
    banner = doc.add_table(rows=1, cols=1)
    banner.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(banner)
    bcell = banner.rows[0].cells[0]
    bcell.width = Inches(6.5)
    set_cell_shading(bcell, COLOR_NAVY_HEX)
    set_cell_margins(bcell, top=120, bottom=120, left=180, right=180)
    bp = bcell.paragraphs[0]
    bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    br = bp.add_run(datos["otec_nombre"].upper())
    br.bold = True
    br.font.size = Pt(11)
    br.font.color.rgb = RGBColor(255, 255, 255)

    # Espacio + Logo centrado
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(24)
    p_sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if logo_bytes:
        p_sp.add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(2.4))

    # Etiqueta tipo documento
    p_tipo = doc.add_paragraph()
    p_tipo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tipo.paragraph_format.space_before = Pt(28)
    p_tipo.paragraph_format.space_after = Pt(6)
    rt = p_tipo.add_run(datos["tipo_documento"].upper())
    rt.bold = True
    rt.font.size = Pt(11)
    rt.font.color.rgb = RGBColor(*COLOR_ACCENT_RGB)

    # Nombre del curso
    p_curso = doc.add_paragraph()
    p_curso.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_curso.paragraph_format.space_before = Pt(4)
    p_curso.paragraph_format.space_after = Pt(10)
    rc = p_curso.add_run(datos["nombre_curso"])
    rc.bold = True
    rc.font.size = Pt(24)
    rc.font.color.rgb = RGBColor(*COLOR_NAVY_RGB)

    # Subtítulo / Módulo / Guía
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(24)
    add_paragraph_border(p_sub, side="bottom", color_hex=COLOR_ACCENT_HEX, sz="12")
    rs = p_sub.add_run(datos["titulo_guia"])
    rs.font.size = Pt(15)
    rs.font.color.rgb = RGBColor(*COLOR_DARK_TEXT_RGB)

    # Espaciador antes de la ficha técnica
    for _ in range(3):
        doc.add_paragraph()

    # Ficha técnica de portada
    meta_rows = [
        ("Relator / Docente", datos["relator"]),
        ("Institución / Cliente", datos["cliente"]),
        ("Modalidad / Duración", f"{datos['modalidad']} — {datos['horas']}"),
        ("Código / Versión", datos["codigo"]),
        ("Fecha", datos["fecha"]),
    ]
    mtbl = doc.add_table(rows=len(meta_rows), cols=2)
    mtbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(mtbl, color_hex=COLOR_BORDER_HEX, sz="4")

    for idx, (label, value) in enumerate(meta_rows):
        c1, c2 = mtbl.rows[idx].cells
        c1.width = Inches(2.2)
        c2.width = Inches(4.3)
        set_cell_shading(c1, COLOR_LIGHT_BG_HEX)
        set_cell_margins(c1, top=90, bottom=90, left=140, right=140)
        set_cell_margins(c2, top=90, bottom=90, left=140, right=140)

        p1 = c1.paragraphs[0]
        r_lbl = p1.add_run(label)
        r_lbl.bold = True
        r_lbl.font.size = Pt(10)
        r_lbl.font.color.rgb = RGBColor(*COLOR_NAVY_RGB)

        p2 = c2.paragraphs[0]
        r_val = p2.add_run(value if value else "Por completar")
        r_val.font.size = Pt(10)
        r_val.font.color.rgb = RGBColor(*COLOR_DARK_TEXT_RGB)

    # Salto de página al contenido
    doc.add_page_break()

    # ------------------------------------------------------------------
    # CUERPO DE LA GUÍA (Plantilla editable para el relator)
    # ------------------------------------------------------------------
    def agregar_seccion(titulo: str, instruccion: str, contenido_inicial: str = ""):
        p_h = doc.add_paragraph()
        p_h.paragraph_format.space_before = Pt(14)
        p_h.paragraph_format.space_after = Pt(6)
        add_paragraph_border(p_h, side="bottom", color_hex=COLOR_ACCENT_HEX, sz="8")
        rh = p_h.add_run(titulo)
        rh.bold = True
        rh.font.size = Pt(13)
        rh.font.color.rgb = RGBColor(*COLOR_NAVY_RGB)

        if contenido_inicial.strip():
            for linea in contenido_inicial.strip().splitlines():
                p_c = doc.add_paragraph(linea)
                p_c.paragraph_format.space_after = Pt(4)
        else:
            # Recuadro guía para el relator
            box = doc.add_table(rows=1, cols=1)
            box.alignment = WD_TABLE_ALIGNMENT.CENTER
            set_table_borders(box, color_hex=COLOR_BORDER_HEX, sz="4")
            bc = box.rows[0].cells[0]
            bc.width = Inches(6.5)
            set_cell_shading(bc, COLOR_LIGHT_BG_HEX)
            set_cell_margins(bc, top=120, bottom=120, left=160, right=160)
            pb = bc.paragraphs[0]
            rb = pb.add_run(instruccion)
            rb.italic = True
            rb.font.size = Pt(10)
            rb.font.color.rgb = RGBColor(*COLOR_MUTED_RGB)

            p_esp = doc.add_paragraph()
            p_esp.paragraph_format.space_after = Pt(12)

    if datos["incluir_secciones"]:
        agregar_seccion(
            "1. Presentación y Objetivos de Aprendizaje",
            "[Espacio para el relator: describir el propósito de la guía y los aprendizajes esperados al finalizar la sesión.]",
            datos["objetivo_general"],
        )
        agregar_seccion(
            "2. Contenidos del Módulo",
            "[Espacio para el relator: enumerar los temas y subtemas que se abordarán en esta guía.]",
            datos["contenidos"],
        )
        agregar_seccion(
            "3. Desarrollo Temático",
            "[Espacio para el relator: redactar el marco conceptual, ejemplos prácticos, procedimientos, esquemas o tablas de apoyo.]",
        )
        agregar_seccion(
            "4. Actividad Práctica / Taller Aplicado",
            "[Espacio para el relator: indicar instrucciones, casos de estudio, preguntas guía, tiempo estimado y criterios de logro.]",
        )
        agregar_seccion(
            "5. Síntesis y Evaluación Formativa",
            "[Espacio para el relator: incluir preguntas de cierre, autoevaluación o lista de cotejo para los participantes.]",
        )
        agregar_seccion(
            "6. Bibliografía y Recursos Complementarios",
            "[Espacio para el relator: citar normativa, manuales, enlaces o lecturas recomendadas.]",
        )
    else:
        p_libre = doc.add_paragraph()
        r_lib = p_libre.add_run("[Espacio disponible para desarrollar el contenido de la guía.]")
        r_lib.italic = True
        r_lib.font.color.rgb = RGBColor(*COLOR_MUTED_RGB)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ==============================================================================
# GENERADOR PDF (.PDF) CON FPDF2
# ==============================================================================
class GuiaOTECPDF(FPDF):
    def __init__(self, datos: dict, logo_bytes: bytes | None):
        super().__init__(orientation="P", unit="mm", format="Letter")
        self.datos = datos
        self.logo_bytes = logo_bytes
        self.set_auto_page_break(auto=True, margin=22)
        self.set_margins(left=22, top=28, right=22)

    def header(self):
        # Sin encabezado en la portada
        if self.page_no() == 1:
            return

        if self.logo_bytes:
            try:
                self.image(io.BytesIO(self.logo_bytes), x=22, y=9, w=26)
            except Exception:
                pass

        self.set_xy(55, 10)
        self.set_font("Helvetica", "B", 9)
        self.set_text_color(*COLOR_NAVY_RGB)
        self.cell(
            139, 5, limpiar_latin1(self.datos["nombre_curso"]), align="R", new_x="LMARGIN", new_y="NEXT"
        )

        self.set_x(55)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(*COLOR_MUTED_RGB)
        sub = f"{self.datos['titulo_guia']}  |  {self.datos['otec_nombre']}"
        self.cell(139, 4, limpiar_latin1(sub), align="R", new_x="LMARGIN", new_y="NEXT")

        # Línea separadora
        self.set_draw_color(*COLOR_ACCENT_RGB)
        self.set_line_width(0.5)
        self.line(22, 22, 194, 22)
        self.set_y(28)

    def footer(self):
        # Sin pie de página en la portada
        if self.page_no() == 1:
            return

        self.set_y(-16)
        self.set_draw_color(203, 213, 224)
        self.set_line_width(0.3)
        self.line(22, self.get_y(), 194, self.get_y())

        self.set_y(-14)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(*COLOR_MUTED_RGB)
        pie = f"{self.datos['otec_nombre']}  |  {self.datos['pie_pagina']}"
        self.cell(135, 6, limpiar_latin1(pie), align="L")

        self.set_font("Helvetica", "B", 8)
        self.set_text_color(*COLOR_NAVY_RGB)
        self.cell(37, 6, limpiar_latin1(f"Página {self.page_no()}"), align="R")


def generar_pdf(datos: dict, logo_bytes: bytes | None) -> bytes:
    pdf = GuiaOTECPDF(datos, logo_bytes)

    # ------------------------------------------------------------------
    # PORTADA
    # ------------------------------------------------------------------
    pdf.add_page()

    # Franja superior
    pdf.set_fill_color(*COLOR_NAVY_RGB)
    pdf.rect(0, 0, 215.9, 22, style="F")
    pdf.set_xy(22, 7)
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(172, 8, limpiar_latin1(datos["otec_nombre"].upper()), align="C")

    # Logo centrado
    if logo_bytes:
        try:
            pdf.image(io.BytesIO(logo_bytes), x=78, y=38, w=60)
            y_cursor = 102
        except Exception:
            y_cursor = 55
    else:
        y_cursor = 55

    # Tipo de documento
    pdf.set_xy(22, y_cursor)
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(*COLOR_ACCENT_RGB)
    pdf.cell(172, 8, limpiar_latin1(datos["tipo_documento"].upper()), align="C", new_x="LMARGIN", new_y="NEXT")

    # Nombre del curso
    pdf.ln(3)
    pdf.set_font("Helvetica", "B", 22)
    pdf.set_text_color(*COLOR_NAVY_RGB)
    pdf.multi_cell(172, 10, limpiar_latin1(datos["nombre_curso"]), align="C")

    # Subtítulo
    pdf.ln(3)
    pdf.set_font("Helvetica", "", 14)
    pdf.set_text_color(*COLOR_DARK_TEXT_RGB)
    pdf.multi_cell(172, 8, limpiar_latin1(datos["titulo_guia"]), align="C")

    # Línea de acento
    pdf.ln(5)
    pdf.set_draw_color(*COLOR_ACCENT_RGB)
    pdf.set_line_width(0.8)
    pdf.line(55, pdf.get_y(), 161, pdf.get_y())

    # Ficha técnica inferior
    meta_rows = [
        ("Relator / Docente", datos["relator"]),
        ("Institución / Cliente", datos["cliente"]),
        ("Modalidad / Duración", f"{datos['modalidad']} - {datos['horas']}"),
        ("Código / Versión", datos["codigo"]),
        ("Fecha", datos["fecha"]),
    ]

    pdf.set_y(188)
    pdf.set_draw_color(203, 213, 224)
    pdf.set_line_width(0.3)

    for label, value in meta_rows:
        pdf.set_x(22)
        pdf.set_fill_color(*COLOR_LIGHT_BG_RGB)
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_text_color(*COLOR_NAVY_RGB)
        pdf.cell(55, 10, limpiar_latin1(f"  {label}"), border=1, fill=True)

        pdf.set_fill_color(255, 255, 255)
        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(*COLOR_DARK_TEXT_RGB)
        val_txt = value if value else "Por completar"
        pdf.cell(117, 10, limpiar_latin1(f"  {val_txt}"), border=1, fill=True, new_x="LMARGIN", new_y="NEXT")

    # Franja inferior decorativa
    pdf.set_fill_color(*COLOR_ACCENT_RGB)
    pdf.rect(0, 272, 215.9, 7.4, style="F")

    # ------------------------------------------------------------------
    # CUERPO DE LA GUÍA
    # ------------------------------------------------------------------
    pdf.add_page()

    def agregar_seccion_pdf(titulo: str, instruccion: str, contenido_inicial: str = ""):
        pdf.set_font("Helvetica", "B", 13)
        pdf.set_text_color(*COLOR_NAVY_RGB)
        pdf.cell(172, 8, limpiar_latin1(titulo), new_x="LMARGIN", new_y="NEXT")

        pdf.set_draw_color(*COLOR_ACCENT_RGB)
        pdf.set_line_width(0.4)
        pdf.line(22, pdf.get_y(), 194, pdf.get_y())
        pdf.ln(4)

        if contenido_inicial.strip():
            pdf.set_font("Helvetica", "", 10.5)
            pdf.set_text_color(*COLOR_DARK_TEXT_RGB)
            pdf.multi_cell(172, 6, limpiar_latin1(contenido_inicial.strip()))
            pdf.ln(6)
        else:
            pdf.set_fill_color(*COLOR_LIGHT_BG_RGB)
            pdf.set_draw_color(203, 213, 224)
            pdf.set_font("Helvetica", "I", 9.5)
            pdf.set_text_color(*COLOR_MUTED_RGB)
            pdf.multi_cell(172, 6.5, limpiar_latin1(instruccion), border=1, fill=True)
            pdf.ln(8)

    if datos["incluir_secciones"]:
        agregar_seccion_pdf(
            "1. Presentación y Objetivos de Aprendizaje",
            "[Espacio para el relator: describir el propósito de la guía y los aprendizajes esperados al finalizar la sesión.]",
            datos["objetivo_general"],
        )
        agregar_seccion_pdf(
            "2. Contenidos del Módulo",
            "[Espacio para el relator: enumerar los temas y subtemas que se abordarán en esta guía.]",
            datos["contenidos"],
        )
        agregar_seccion_pdf(
            "3. Desarrollo Temático",
            "[Espacio para el relator: redactar el marco conceptual, ejemplos prácticos, procedimientos, esquemas o tablas de apoyo.]",
        )
        agregar_seccion_pdf(
            "4. Actividad Práctica / Taller Aplicado",
            "[Espacio para el relator: indicar instrucciones, casos de estudio, preguntas guía, tiempo estimado y criterios de logro.]",
        )
        agregar_seccion_pdf(
            "5. Síntesis y Evaluación Formativa",
            "[Espacio para el relator: incluir preguntas de cierre, autoevaluación o lista de cotejo para los participantes.]",
        )
        agregar_seccion_pdf(
            "6. Bibliografía y Recursos Complementarios",
            "[Espacio para el relator: citar normativa, manuales, enlaces o lecturas recomendadas.]",
        )
    else:
        pdf.set_font("Helvetica", "I", 10)
        pdf.set_text_color(*COLOR_MUTED_RGB)
        pdf.multi_cell(172, 7, limpiar_latin1("[Espacio disponible para desarrollar el contenido de la guía.]"))

    return bytes(pdf.output())


# ==============================================================================
# INTERFAZ STREAMLIT (LISTA PARA IMPORTAR CON run())
# ==============================================================================
def run():
    st.title("Generador de Guías de Capacitación — OTEC Santa Mónica")
    st.caption(
        "Genera plantillas profesionales en Word (.docx) y PDF (.pdf) con portada, "
        "encabezado y pie de página institucional para relatores."
    )

    # Carga de logo desde GitHub con respaldo manual (útil si el repo es privado)
    logo_bytes = descargar_logo(LOGO_URL_DEFAULT)

    with st.expander("Configuración del Logo Institucional", expanded=(logo_bytes is None)):
        if logo_bytes:
            col_l1, col_l2 = st.columns([1, 3])
            with col_l1:
                st.image(logo_bytes, width=150)
            with col_l2:
                st.success("Logo cargado automáticamente desde GitHub.")
        else:
            st.warning(
                "No se pudo descargar el logo desde la URL de GitHub (verifica si el repositorio es privado). "
                "Puedes subir el archivo JPG/PNG manualmente aquí:"
            )

        logo_subido = st.file_uploader(
            "Reemplazar o cargar logo manualmente (opcional)",
            type=["jpg", "jpeg", "png"],
        )
        if logo_subido is not None:
            logo_bytes = logo_subido.read()
            st.image(logo_bytes, width=150)

    # Formulario principal
    st.subheader("Datos de Portada y Encabezado")
    col1, col2 = st.columns(2)

    with col1:
        otec_nombre = st.text_input(
            "Nombre OTEC",
            value="Centro de Capacitación Santa Mónica Limitada",
        )
        tipo_documento = st.selectbox(
            "Tipo de Documento",
            [
                "Guía de Capacitación y Relatoría",
                "Manual del Participante",
                "Guía de Taller Práctico",
                "Material de Apoyo Docente",
            ],
        )
        nombre_curso = st.text_input(
            "Nombre del Curso",
            value="Inteligencia Artificial y Análisis de Datos Aplicado",
        )
        titulo_guia = st.text_input(
            "Subtítulo / Módulo / Guía",
            value="Módulo 1: Fundamentos y Herramientas Prácticas",
        )

    with col2:
        relator = st.text_input("Nombre del Relator / Docente", value="")
        cliente = st.text_input("Institución / Cliente", value="")
        c_mod, c_hrs = st.columns(2)
        with c_mod:
            modalidad = st.selectbox("Modalidad", ["Presencial", "Online Sincrónico", "E-Learning", "Mixta"])
        with c_hrs:
            horas = st.text_input("Duración", value="16 horas cronológicas")
        c_cod, c_fec = st.columns(2)
        with c_cod:
            codigo = st.text_input("Código / Versión", value="GUIA-01 / v1.0")
        with c_fec:
            fecha_sel = st.date_input("Fecha", value=date.today())

    pie_pagina = st.text_input(
        "Texto de Pie de Página",
        value="Material exclusivo de capacitación — Prohibida su reproducción sin autorización",
    )

    st.subheader("Estructura Base para el Relator")
    incluir_secciones = st.checkbox(
        "Incluir secciones guía (Objetivos, Contenidos, Desarrollo, Taller, Evaluación, Bibliografía)",
        value=True,
    )

    objetivo_general = ""
    contenidos = ""
    if incluir_secciones:
        with st.expander("Completar Objetivo y Contenidos ahora (opcional)"):
            objetivo_general = st.text_area(
                "Objetivo General (si se deja vacío, aparecerá el recuadro editable para el relator)",
                height=90,
            )
            contenidos = st.text_area(
                "Contenidos del Módulo (opcional)",
                height=100,
            )

    datos = {
        "otec_nombre": otec_nombre,
        "tipo_documento": tipo_documento,
        "nombre_curso": nombre_curso,
        "titulo_guia": titulo_guia,
        "relator": relator,
        "cliente": cliente,
        "modalidad": modalidad,
        "horas": horas,
        "codigo": codigo,
        "fecha": fecha_sel.strftime("%d/%m/%Y"),
        "pie_pagina": pie_pagina,
        "incluir_secciones": incluir_secciones,
        "objetivo_general": objetivo_general,
        "contenidos": contenidos,
    }

    st.divider()

    if st.button("Generar Documentos (Word y PDF)", type="primary", use_container_width=True):
        if not nombre_curso.strip():
            st.error("Ingresa el nombre del curso antes de generar los documentos.")
            return

        with st.spinner("Generando plantilla profesional en Word y PDF..."):
            docx_bytes = generar_word(datos, logo_bytes)
            pdf_bytes = generar_pdf(datos, logo_bytes)

        st.session_state["docx_generado"] = docx_bytes
        st.session_state["pdf_generado"] = pdf_bytes
        nombre_base = "".join(c if c.isalnum() or c in (" ", "-", "_") else "" for c in nombre_curso).strip()
        nombre_base = nombre_base.replace(" ", "_") or "Guia_OTEC_Santa_Monica"
        st.session_state["nombre_archivo"] = f"Guia_{nombre_base}"
        st.success("Documentos generados correctamente. Descárgalos abajo.")

    if "docx_generado" in st.session_state and "pdf_generado" in st.session_state:
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.download_button(
                label="Descargar Word (.docx)",
                data=st.session_state["docx_generado"],
                file_name=f"{st.session_state['nombre_archivo']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with col_d2:
            st.download_button(
                label="Descargar PDF (.pdf)",
                data=st.session_state["pdf_generado"],
                file_name=f"{st.session_state['nombre_archivo']}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )


if __name__ == "__main__":
    st.set_page_config(page_title="Generador de Guías OTEC", layout="wide")
    run()
