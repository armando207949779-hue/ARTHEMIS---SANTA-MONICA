import io
import os
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

LOGO_LOCAL_PATH = "OTEC_Santa_Monica_SII.jpg"
LOGO_URL_DEFAULT = (
    "https://raw.githubusercontent.com/armando207949779-hue/"
    "ARTHEMIS---SANTA-MONICA/main/OTEC_Santa_Monica_SII.jpg"
)

# Paleta corporativa por defecto
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
def cargar_logo_default() -> bytes | None:
    """1) Intenta leer el logo local en el repo. 2) Si no existe, lo descarga de GitHub."""
    if os.path.exists(LOGO_LOCAL_PATH):
        try:
            with open(LOGO_LOCAL_PATH, "rb") as f:
                return f.read()
        except Exception:
            pass
    try:
        resp = requests.get(LOGO_URL_DEFAULT, timeout=10)
        if resp.status_code == 200 and resp.content:
            return resp.content
    except Exception:
        pass
    return None


def limpiar_latin1(texto: str) -> str:
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


def truncar_texto(texto: str, max_chars: int) -> str:
    if len(texto) <= max_chars:
        return texto
    return texto[: max_chars - 3].rstrip() + "..."


# ==============================================================================
# UTILIDADES XML PARA PYTHON-DOCX
# ==============================================================================
def set_cell_shading(cell, color_hex: str):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_margins(cell, top=110, bottom=110, left=150, right=150):
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
        f'  <w:left w:val="none"/><w:right w:val="none"/><w:insideV w:val="none"/>'
        f"</w:tblBorders>"
    )
    tblPr.append(borders)


def set_box_borders(table, color_hex=COLOR_BORDER_HEX, accent_left=COLOR_ACCENT_HEX):
    """Recuadro limpio con borde izquierdo de acento corporativo."""
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="{color_hex}"/>'
        f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="{color_hex}"/>'
        f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="{color_hex}"/>'
        f'  <w:left w:val="single" w:sz="18" w:space="0" w:color="{accent_left}"/>'
        f'  <w:insideH w:val="none"/><w:insideV w:val="none"/>'
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


def add_xml_field(run, field_name: str):
    """Inserta campos dinámicos de Word como PAGE o NUMPAGES."""
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = field_name
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

    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(0.9)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    section.different_first_page_header_footer = True

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
    c_logo.width = Inches(1.6)
    c_info.width = Inches(4.9)

    p_hlogo = c_logo.paragraphs[0]
    p_hlogo.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if logo_bytes:
        p_hlogo.add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(1.1))
    else:
        r = p_hlogo.add_run("OTEC SANTA MÓNICA")
        r.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(*COLOR_NAVY_RGB)

    p_hinfo = c_info.paragraphs[0]
    p_hinfo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r1 = p_hinfo.add_run(f"{truncar_texto(datos['nombre_curso'], 65)}\n")
    r1.bold = True
    r1.font.size = Pt(9)
    r1.font.color.rgb = RGBColor(*COLOR_NAVY_RGB)
    r2 = p_hinfo.add_run(f"{truncar_texto(datos['titulo_guia'], 55)}  |  OTEC Santa Mónica")
    r2.font.size = Pt(8)
    r2.font.color.rgb = RGBColor(*COLOR_MUTED_RGB)

    p_hline = header.add_paragraph()
    p_hline.paragraph_format.space_before = Pt(2)
    p_hline.paragraph_format.space_after = Pt(6)
    add_paragraph_border(p_hline, side="bottom", color_hex=COLOR_ACCENT_HEX, sz="8")

    # ------------------------------------------------------------------
    # PIE DE PÁGINA (Páginas 2 en adelante — "Página X de Y")
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
    fc1.width = Inches(5.0)
    fc2.width = Inches(1.5)

    p_f1 = fc1.paragraphs[0]
    rf1 = p_f1.add_run(f"OTEC Santa Mónica  |  {datos['pie_pagina']}")
    rf1.font.size = Pt(8)
    rf1.font.color.rgb = RGBColor(*COLOR_MUTED_RGB)

    p_f2 = fc2.paragraphs[0]
    p_f2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for txt, is_field in [("Página ", None), ("", "PAGE"), (" de ", None), ("", "NUMPAGES")]:
        r = p_f2.add_run(txt)
        r.font.size = Pt(8)
        if is_field:
            r.bold = True
            r.font.color.rgb = RGBColor(*COLOR_NAVY_RGB)
            add_xml_field(r, is_field)
        else:
            r.font.color.rgb = RGBColor(*COLOR_MUTED_RGB)

    # ------------------------------------------------------------------
    # PORTADA (Página 1)
    # ------------------------------------------------------------------
    banner = doc.add_table(rows=1, cols=1)
    banner.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(banner)
    bcell = banner.rows[0].cells[0]
    bcell.width = Inches(6.5)
    set_cell_shading(bcell, COLOR_NAVY_HEX)
    set_cell_margins(bcell, top=130, bottom=130, left=180, right=180)
    bp = bcell.paragraphs[0]
    bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    br = bp.add_run(datos["otec_nombre"].upper())
    br.bold = True
    br.font.size = Pt(11)
    br.font.color.rgb = RGBColor(255, 255, 255)

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(22)
    p_sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if logo_bytes:
        p_sp.add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(2.3))

    p_tipo = doc.add_paragraph()
    p_tipo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tipo.paragraph_format.space_before = Pt(24)
    p_tipo.paragraph_format.space_after = Pt(6)
    rt = p_tipo.add_run(datos["tipo_documento"].upper())
    rt.bold = True
    rt.font.size = Pt(11)
    rt.font.color.rgb = RGBColor(*COLOR_ACCENT_RGB)

    p_curso = doc.add_paragraph()
    p_curso.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_curso.paragraph_format.space_before = Pt(4)
    p_curso.paragraph_format.space_after = Pt(10)
    rc = p_curso.add_run(datos["nombre_curso"])
    rc.bold = True
    rc.font.size = Pt(24)
    rc.font.color.rgb = RGBColor(*COLOR_NAVY_RGB)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(22)
    add_paragraph_border(p_sub, side="bottom", color_hex=COLOR_ACCENT_HEX, sz="12")
    rs = p_sub.add_run(datos["titulo_guia"])
    rs.font.size = Pt(15)
    rs.font.color.rgb = RGBColor(*COLOR_DARK_TEXT_RGB)

    for _ in range(3):
        doc.add_paragraph()

    placeholder_vacio = "____________________________" if datos["linea_firma_vacia"] else "Por completar"
    meta_rows = [
        ("Relator / Docente", datos["relator"] or placeholder_vacio),
        ("Institución / Cliente", datos["cliente"] or placeholder_vacio),
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
        set_cell_margins(c1, top=95, bottom=95, left=140, right=140)
        set_cell_margins(c2, top=95, bottom=95, left=140, right=140)

        p1 = c1.paragraphs[0]
        r_lbl = p1.add_run(label)
        r_lbl.bold = True
        r_lbl.font.size = Pt(10)
        r_lbl.font.color.rgb = RGBColor(*COLOR_NAVY_RGB)

        p2 = c2.paragraphs[0]
        r_val = p2.add_run(value)
        r_val.font.size = Pt(10)
        r_val.font.color.rgb = RGBColor(*COLOR_DARK_TEXT_RGB)

    doc.add_page_break()

    # ------------------------------------------------------------------
    # CUERPO DE LA GUÍA
    # ------------------------------------------------------------------
    def agregar_seccion(titulo: str, instruccion: str, contenido_inicial: str = "", espacios_extra: int = 1):
        p_h = doc.add_paragraph()
        p_h.paragraph_format.space_before = Pt(14)
        p_h.paragraph_format.space_after = Pt(6)
        p_h.paragraph_format.keep_with_next = True
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
            box = doc.add_table(rows=1, cols=1)
            box.alignment = WD_TABLE_ALIGNMENT.CENTER
            set_box_borders(box, color_hex=COLOR_BORDER_HEX, accent_left=COLOR_ACCENT_HEX)
            bc = box.rows[0].cells[0]
            bc.width = Inches(6.5)
            set_cell_shading(bc, COLOR_LIGHT_BG_HEX)
            set_cell_margins(bc, top=140, bottom=140, left=180, right=160)
            pb = bc.paragraphs[0]
            rb = pb.add_run(instruccion)
            rb.italic = True
            rb.font.size = Pt(10)
            rb.font.color.rgb = RGBColor(*COLOR_MUTED_RGB)

            for _ in range(espacios_extra):
                p_esp = doc.add_paragraph()
                p_esp.paragraph_format.space_after = Pt(10)

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
            espacios_extra=2,
        )
        agregar_seccion(
            "4. Actividad Práctica / Taller Aplicado",
            "[Espacio para el relator: indicar instrucciones, casos de estudio, preguntas guía, tiempo estimado y criterios de logro.]",
            espacios_extra=2,
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

    # Bloque opcional de Control del Documento
    if datos["incluir_control_doc"]:
        p_ctrl = doc.add_paragraph()
        p_ctrl.paragraph_format.space_before = Pt(18)
        p_ctrl.paragraph_format.space_after = Pt(6)
        p_ctrl.paragraph_format.keep_with_next = True
        r_ct = p_ctrl.add_run("Control y Validación del Documento")
        r_ct.bold = True
        r_ct.font.size = Pt(11)
        r_ct.font.color.rgb = RGBColor(*COLOR_NAVY_RGB)

        ctbl = doc.add_table(rows=2, cols=3)
        ctbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(ctbl, color_hex=COLOR_BORDER_HEX, sz="4")
        encabezados = ["Elaborado por (Relator)", "Revisado por (Coordinación OTEC)", "Fecha de Aprobación"]
        valores = [datos["relator"] or "Por completar", "Dirección Académica", datos["fecha"]]
        for i in range(3):
            c_top = ctbl.rows[0].cells[i]
            c_bot = ctbl.rows[1].cells[i]
            set_cell_shading(c_top, COLOR_LIGHT_BG_HEX)
            set_cell_margins(c_top, top=80, bottom=80, left=100, right=100)
            set_cell_margins(c_bot, top=100, bottom=100, left=100, right=100)
            rt_c = c_top.paragraphs[0].add_run(encabezados[i])
            rt_c.bold = True
            rt_c.font.size = Pt(9)
            rt_c.font.color.rgb = RGBColor(*COLOR_NAVY_RGB)
            rb_c = c_bot.paragraphs[0].add_run(valores[i])
            rb_c.font.size = Pt(9)

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
        self.alias_nb_pages()  # Habilita {nb} para "Página X de Y"
        self.set_auto_page_break(auto=True, margin=24)
        self.set_margins(left=22, top=28, right=22)

    def header(self):
        if self.page_no() == 1:
            return

        if self.logo_bytes:
            try:
                self.image(io.BytesIO(self.logo_bytes), x=22, y=9, w=24)
            except Exception:
                pass

        self.set_xy(52, 10)
        self.set_font("Helvetica", "B", 9)
        self.set_text_color(*COLOR_NAVY_RGB)
        curso_hdr = truncar_texto(self.datos["nombre_curso"], 62)
        self.cell(142, 5, limpiar_latin1(curso_hdr), align="R", new_x="LMARGIN", new_y="NEXT")

        self.set_x(52)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(*COLOR_MUTED_RGB)
        sub = f"{truncar_texto(self.datos['titulo_guia'], 50)}  |  OTEC Santa Mónica"
        self.cell(142, 4, limpiar_latin1(sub), align="R", new_x="LMARGIN", new_y="NEXT")

        self.set_draw_color(*COLOR_ACCENT_RGB)
        self.set_line_width(0.5)
        self.line(22, 22, 194, 22)
        self.set_y(28)

    def footer(self):
        if self.page_no() == 1:
            return

        self.set_y(-17)
        self.set_draw_color(203, 213, 224)
        self.set_line_width(0.3)
        self.line(22, self.get_y(), 194, self.get_y())

        self.set_y(-14.5)
        self.set_font("Helvetica", "", 7.5)
        self.set_text_color(*COLOR_MUTED_RGB)
        pie = truncar_texto(f"OTEC Santa Mónica  |  {self.datos['pie_pagina']}", 88)
        self.cell(138, 6, limpiar_latin1(pie), align="L")

        self.set_font("Helvetica", "B", 8)
        self.set_text_color(*COLOR_NAVY_RGB)
        self.cell(34, 6, limpiar_latin1(f"Página {self.page_no()} de {{nb}}"), align="R")


def generar_pdf(datos: dict, logo_bytes: bytes | None) -> bytes:
    pdf = GuiaOTECPDF(datos, logo_bytes)

    # ------------------------------------------------------------------
    # PORTADA
    # ------------------------------------------------------------------
    pdf.add_page()

    pdf.set_fill_color(*COLOR_NAVY_RGB)
    pdf.rect(0, 0, 215.9, 22, style="F")
    pdf.set_xy(22, 7)
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(172, 8, limpiar_latin1(datos["otec_nombre"].upper()), align="C")

    if logo_bytes:
        try:
            pdf.image(io.BytesIO(logo_bytes), x=78, y=36, w=60)
            y_cursor = 100
        except Exception:
            y_cursor = 55
    else:
        y_cursor = 55

    pdf.set_xy(22, y_cursor)
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(*COLOR_ACCENT_RGB)
    pdf.cell(172, 8, limpiar_latin1(datos["tipo_documento"].upper()), align="C", new_x="LMARGIN", new_y="NEXT")

    pdf.ln(3)
    pdf.set_font("Helvetica", "B", 22)
    pdf.set_text_color(*COLOR_NAVY_RGB)
    pdf.multi_cell(172, 10, limpiar_latin1(datos["nombre_curso"]), align="C")

    pdf.ln(3)
    pdf.set_font("Helvetica", "", 14)
    pdf.set_text_color(*COLOR_DARK_TEXT_RGB)
    pdf.multi_cell(172, 8, limpiar_latin1(datos["titulo_guia"]), align="C")

    pdf.ln(5)
    pdf.set_draw_color(*COLOR_ACCENT_RGB)
    pdf.set_line_width(0.8)
    pdf.line(55, pdf.get_y(), 161, pdf.get_y())

    placeholder_vacio = "____________________________" if datos["linea_firma_vacia"] else "Por completar"
    meta_rows = [
        ("Relator / Docente", datos["relator"] or placeholder_vacio),
        ("Institución / Cliente", datos["cliente"] or placeholder_vacio),
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
        val_txt = truncar_texto(value, 54)
        pdf.cell(117, 10, limpiar_latin1(f"  {val_txt}"), border=1, fill=True, new_x="LMARGIN", new_y="NEXT")

    pdf.set_fill_color(*COLOR_ACCENT_RGB)
    pdf.rect(0, 272, 215.9, 7.4, style="F")

    # ------------------------------------------------------------------
    # CUERPO DE LA GUÍA
    # ------------------------------------------------------------------
    pdf.add_page()

    def agregar_seccion_pdf(titulo: str, instruccion: str, contenido_inicial: str = "", altura_extra: int = 0):
        # Evita títulos huérfanos al final de la página
        if pdf.get_y() > 225:
            pdf.add_page()

        pdf.set_font("Helvetica", "B", 13)
        pdf.set_text_color(*COLOR_NAVY_RGB)
        pdf.cell(172, 8, limpiar_latin1(titulo), new_x="LMARGIN", new_y="NEXT")

        pdf.set_draw_color(*COLOR_ACCENT_RGB)
        pdf.set_line_width(0.45)
        pdf.line(22, pdf.get_y(), 194, pdf.get_y())
        pdf.ln(4)

        if contenido_inicial.strip():
            pdf.set_font("Helvetica", "", 10.5)
            pdf.set_text_color(*COLOR_DARK_TEXT_RGB)
            pdf.multi_cell(172, 6, limpiar_latin1(contenido_inicial.strip()))
            pdf.ln(7)
        else:
            # Recuadro con padding interno y barra lateral de acento
            y_inicio = pdf.get_y()
            pdf.set_fill_color(*COLOR_LIGHT_BG_RGB)
            pdf.set_draw_color(203, 213, 224)
            pdf.set_line_width(0.3)

            pdf.set_xy(26, y_inicio + 3)
            pdf.set_font("Helvetica", "I", 9.5)
            pdf.set_text_color(*COLOR_MUTED_RGB)
            # Medimos altura necesaria simulando en un bloque con relleno
            pdf.set_xy(22, y_inicio)
            pdf.multi_cell(172, 6.5, limpiar_latin1(f"  {instruccion}"), border=1, fill=True, padding=(3, 4, 3 + altura_extra, 4))
            y_fin = pdf.get_y()

            # Barra de acento a la izquierda del recuadro
            pdf.set_fill_color(*COLOR_ACCENT_RGB)
            pdf.rect(22, y_inicio, 1.5, y_fin - y_inicio, style="F")
            pdf.set_y(y_fin + 7)

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
            altura_extra=12,
        )
        agregar_seccion_pdf(
            "4. Actividad Práctica / Taller Aplicado",
            "[Espacio para el relator: indicar instrucciones, casos de estudio, preguntas guía, tiempo estimado y criterios de logro.]",
            altura_extra=12,
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

    # Bloque opcional de Control del Documento
    if datos["incluir_control_doc"]:
        if pdf.get_y() > 230:
            pdf.add_page()
        pdf.ln(4)
        pdf.set_font("Helvetica", "B", 10.5)
        pdf.set_text_color(*COLOR_NAVY_RGB)
        pdf.cell(172, 7, limpiar_latin1("Control y Validación del Documento"), new_x="LMARGIN", new_y="NEXT")

        anchos = [62, 66, 44]
        encabezados = ["Elaborado por (Relator)", "Revisado por (OTEC)", "Fecha Aprobación"]
        valores = [truncar_texto(datos["relator"] or "Por completar", 28), "Dirección Académica", datos["fecha"]]

        pdf.set_draw_color(203, 213, 224)
        pdf.set_fill_color(*COLOR_LIGHT_BG_RGB)
        pdf.set_font("Helvetica", "B", 8.5)
        for w, h in zip(anchos, encabezados):
            pdf.cell(w, 8, limpiar_latin1(f" {h}"), border=1, fill=True)
        pdf.ln()

        pdf.set_fill_color(255, 255, 255)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.set_text_color(*COLOR_DARK_TEXT_RGB)
        for w, v in zip(anchos, valores):
            pdf.cell(w, 9, limpiar_latin1(f" {v}"), border=1, fill=True)
        pdf.ln()

    return bytes(pdf.output())


# ==============================================================================
# INTERFAZ STREAMLIT (LISTA PARA IMPORTAR CON run())
# ==============================================================================
def run():
    st.title("Generador de Guías de Capacitación — OTEC Santa Mónica")
    st.caption(
        "Genera plantillas profesionales en Word (.docx) y PDF (.pdf) con portada, "
        "encabezado, pie de página con numeración (Página X de Y) y control documental."
    )

    logo_bytes = cargar_logo_default()

    with st.expander("Configuración del Logo Institucional", expanded=(logo_bytes is None)):
        if logo_bytes:
            col_l1, col_l2 = st.columns([1, 3])
            with col_l1:
                st.image(logo_bytes, width=150)
            with col_l2:
                st.success("Logo institucional cargado correctamente.")
        else:
            st.warning("No se encontró el logo automáticamente. Sube el archivo JPG/PNG aquí:")

        logo_subido = st.file_uploader("Reemplazar logo (opcional)", type=["jpg", "jpeg", "png"])
        if logo_subido is not None:
            logo_bytes = logo_subido.read()
            st.image(logo_bytes, width=150)

    st.subheader("Datos de Portada y Encabezado")
    col1, col2 = st.columns(2)

    with col1:
        otec_nombre = st.text_input("Nombre OTEC", value="Centro de Capacitación Santa Mónica Limitada")
        tipo_documento = st.selectbox(
            "Tipo de Documento",
            [
                "Guía de Capacitación y Relatoría",
                "Manual del Participante",
                "Guía de Taller Práctico",
                "Material de Apoyo Docente",
            ],
        )
        nombre_curso = st.text_input("Nombre del Curso", value="Inteligencia Artificial y Análisis de Datos Aplicado")
        titulo_guia = st.text_input("Subtítulo / Módulo / Guía", value="Módulo 1: Fundamentos y Herramientas Prácticas")

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
        "Texto de Pie de Página (breve para evitar cortes)",
        value="Material exclusivo de capacitación — Prohibida su reproducción",
    )

    st.subheader("Estructura y Opciones")
    c_op1, c_op2, c_op3 = st.columns(3)
    with c_op1:
        incluir_secciones = st.checkbox("Incluir 6 secciones guía", value=True)
    with c_op2:
        incluir_control_doc = st.checkbox("Incluir tabla de Control Documental", value=True)
    with c_op3:
        linea_firma_vacia = st.checkbox("Usar línea (____) en campos vacíos", value=False)

    objetivo_general = ""
    contenidos = ""
    if incluir_secciones:
        with st.expander("Completar Objetivo y Contenidos ahora (opcional)"):
            objetivo_general = st.text_area("Objetivo General (opcional)", height=90)
            contenidos = st.text_area("Contenidos del Módulo (opcional)", height=100)

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
        "incluir_control_doc": incluir_control_doc,
        "linea_firma_vacia": linea_firma_vacia,
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
