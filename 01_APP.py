import io
import os
from datetime import date
import requests
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from fpdf import FPDF

LOGO_LOCAL_PATH = "OTEC_Santa_Monica_SII.jpg"
LOGO_URL_DEFAULT = (
    "https://raw.githubusercontent.com/armando207949779-hue/"
    "ARTHEMIS---SANTA-MONICA/main/OTEC_Santa_Monica_SII.jpg"
)

PALETAS = {
    "Azul Institucional (Clásico)": ("#1B365D", "#2B6CB0"),
    "Azul Petróleo Moderno": ("#0F3D3E", "#1F8A70"),
    "Grafito y Azul Acero": ("#1F2937", "#3B82F6"),
    "Burdeo Corporativo": ("#4A1525", "#9B2C2C"),
}

COLOR_DARK_TEXT_RGB = (38, 46, 58)
COLOR_MUTED_RGB = (113, 128, 150)
COLOR_BORDER_HEX = "D8E0EA"
COLOR_BORDER_RGB = (216, 224, 234)
COLOR_SOFT_BG_HEX = "F7F9FC"
COLOR_SOFT_BG_RGB = (247, 249, 252)


# ==============================================================================
# UTILIDADES GENERALES
# ==============================================================================
def hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    h = hex_color.lstrip("#")
    return tuple(int(h[i : i + 2], 16) for i in (0, 2, 4))


def hex_clean(hex_color: str) -> str:
    return hex_color.lstrip("#").upper()


@st.cache_data(show_spinner=False)
def cargar_logo_default() -> bytes | None:
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
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_clean(color_hex)}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_row_height(row, inches_val: float):
    trPr = row._tr.get_or_add_trPr()
    trHeight = OxmlElement("w:trHeight")
    trHeight.set(qn("w:val"), str(int(inches_val * 1440)))
    trHeight.set(qn("w:hRule"), "atLeast")
    trPr.append(trHeight)


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


def set_cell_border(cell, **kwargs):
    """Ejemplo: set_cell_border(cell, right={"val": "single", "sz": "24", "color": "2B6CB0"})"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    for edge, edge_data in kwargs.items():
        tag = f"w:{edge}"
        element = tcBorders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            tcBorders.append(element)
        for key, value in edge_data.items():
            element.set(qn(f"w:{key}"), str(value))


def add_paragraph_border(paragraph, side="bottom", color_hex="2B6CB0", sz="8"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    border = OxmlElement(f"w:{side}")
    border.set(qn("w:val"), "single")
    border.set(qn("w:sz"), sz)
    border.set(qn("w:space"), "4")
    border.set(qn("w:color"), hex_clean(color_hex))
    pBdr.append(border)
    pPr.append(pBdr)


def add_xml_field(run, field_name: str):
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
# GENERADOR WORD (.DOCX) — SOLO PORTADA + HOJA(S) LIMPIA(S) DE FORMATO
# ==============================================================================
def generar_word(datos: dict, logo_bytes: bytes | None) -> bytes:
    doc = Document()

    primary_hex = hex_clean(datos["color_primario"])
    accent_hex = hex_clean(datos["color_acento"])
    primary_rgb = hex_to_rgb(datos["color_primario"])
    accent_rgb = hex_to_rgb(datos["color_acento"])

    # Estilos corporativos precargados para que el relator escriba con formato profesional
    style_normal = doc.styles["Normal"]
    style_normal.font.name = "Calibri"
    style_normal.font.size = Pt(11)
    style_normal.font.color.rgb = RGBColor(*COLOR_DARK_TEXT_RGB)

    if "Heading 1" in doc.styles:
        h1 = doc.styles["Heading 1"]
        h1.font.name = "Calibri"
        h1.font.size = Pt(15)
        h1.font.bold = True
        h1.font.color.rgb = RGBColor(*primary_rgb)

    if "Heading 2" in doc.styles:
        h2 = doc.styles["Heading 2"]
        h2.font.name = "Calibri"
        h2.font.size = Pt(12.5)
        h2.font.bold = True
        h2.font.color.rgb = RGBColor(*accent_rgb)

    # ------------------------------------------------------------------
    # SECCIÓN 1: PORTADA EDITORIAL (Diseño asimétrico con barra lateral)
    # ------------------------------------------------------------------
    sec1 = doc.sections[0]
    sec1.page_width = Inches(8.5)
    sec1.page_height = Inches(11.0)
    sec1.top_margin = Inches(0.55)
    sec1.bottom_margin = Inches(0.55)
    sec1.left_margin = Inches(0.55)
    sec1.right_margin = Inches(0.75)

    # Tabla maestra de 2 columnas: [Franja lateral | Lienzo principal de portada]
    cover_tbl = doc.add_table(rows=1, cols=2)
    cover_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cover_tbl.autofit = False
    remove_table_borders(cover_tbl)
    set_row_height(cover_tbl.rows[0], 9.75)

    c_bar, c_main = cover_tbl.rows[0].cells
    c_bar.width = Inches(0.32)
    c_main.width = Inches(6.88)

    # Barra lateral izquierda (Primario + borde derecho en Acento)
    set_cell_shading(c_bar, primary_hex)
    set_cell_border(c_bar, right={"val": "single", "sz": "24", "space": "0", "color": accent_hex})
    c_bar.paragraphs[0].text = ""

    # Contenido principal de la portada
    set_cell_margins(c_main, top=180, bottom=140, left=420, right=120)
    c_main.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    # 1) Cabecera de portada: Logo izquierda + Código/Versión derecha
    top_tbl = c_main.add_table(rows=1, cols=2)
    top_tbl.autofit = False
    remove_table_borders(top_tbl)
    tc_logo, tc_code = top_tbl.rows[0].cells
    tc_logo.width = Inches(3.8)
    tc_code.width = Inches(2.5)

    p_logo = tc_logo.paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if logo_bytes:
        p_logo.add_run().add_picture(io.BytesIO(logo_bytes), height=Inches(0.95))
    else:
        r_lg = p_logo.add_run("OTEC SANTA MÓNICA")
        r_lg.bold = True
        r_lg.font.size = Pt(14)
        r_lg.font.color.rgb = RGBColor(*primary_rgb)

    p_code = tc_code.paragraphs[0]
    p_code.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_cod1 = p_code.add_run(f"{datos['codigo']}\n")
    r_cod1.bold = True
    r_cod1.font.size = Pt(9.5)
    r_cod1.font.color.rgb = RGBColor(*primary_rgb)
    r_cod2 = p_code.add_run(datos["fecha"])
    r_cod2.font.size = Pt(9)
    r_cod2.font.color.rgb = RGBColor(*COLOR_MUTED_RGB)

    # Limpiamos párrafo vacío inicial de la celda c_main si quedó arriba
    c_main.paragraphs[0].paragraph_format.space_before = Pt(0)
    c_main.paragraphs[0].paragraph_format.space_after = Pt(0)

    # 2) Bloque central: Tipo de documento + Título Curso + Subtítulo
    p_tipo = c_main.add_paragraph()
    p_tipo.paragraph_format.space_before = Pt(90)
    p_tipo.paragraph_format.space_after = Pt(8)
    rt = p_tipo.add_run(datos["tipo_documento"].upper())
    rt.bold = True
    rt.font.size = Pt(10.5)
    rt.font.color.rgb = RGBColor(*accent_rgb)

    p_curso = c_main.add_paragraph()
    p_curso.paragraph_format.space_before = Pt(0)
    p_curso.paragraph_format.space_after = Pt(12)
    p_curso.paragraph_format.line_spacing = 1.08
    rc = p_curso.add_run(datos["nombre_curso"])
    rc.bold = True
    rc.font.size = Pt(27)
    rc.font.color.rgb = RGBColor(*primary_rgb)

    p_sub = c_main.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(18)
    rs = p_sub.add_run(datos["titulo_guia"])
    rs.font.size = Pt(14.5)
    rs.font.color.rgb = RGBColor(*COLOR_MUTED_RGB)

    # Línea de acento corta bajo el subtítulo
    rule_tbl = c_main.add_table(rows=1, cols=2)
    rule_tbl.autofit = False
    remove_table_borders(rule_tbl)
    rc1, rc2 = rule_tbl.rows[0].cells
    rc1.width = Inches(1.5)
    rc2.width = Inches(4.8)
    set_cell_border(rc1, top={"val": "single", "sz": "24", "space": "0", "color": accent_hex})
    rc1.paragraphs[0].paragraph_format.space_after = Pt(0)
    rc2.paragraphs[0].paragraph_format.space_after = Pt(0)

    # Espaciador hacia la ficha técnica inferior
    p_gap = c_main.add_paragraph()
    p_gap.paragraph_format.space_before = Pt(105)
    p_gap.paragraph_format.space_after = Pt(6)

    # 3) Ficha técnica minimalista (2x2 con fondo suave)
    placeholder = "______________________" if datos["linea_vacia"] else "Por completar"
    ficha_items = [
        ("RELATOR / DOCENTE", datos["relator"] or placeholder),
        ("INSTITUCIÓN / CLIENTE", datos["cliente"] or placeholder),
        ("MODALIDAD Y DURACIÓN", f"{datos['modalidad']}  ·  {datos['horas']}"),
        ("ORGANISMO TÉCNICO", datos["otec_nombre"]),
    ]

    meta_tbl = c_main.add_table(rows=2, cols=2)
    meta_tbl.autofit = False
    remove_table_borders(meta_tbl)

    for idx, (lbl, val) in enumerate(ficha_items):
        r_idx, c_idx = divmod(idx, 2)
        cell = meta_tbl.rows[r_idx].cells[c_idx]
        cell.width = Inches(3.15)
        set_cell_shading(cell, COLOR_SOFT_BG_HEX)
        set_cell_margins(cell, top=110, bottom=110, left=140, right=120)
        set_cell_border(
            cell,
            top={"val": "single", "sz": "4", "space": "0", "color": COLOR_BORDER_HEX},
            bottom={"val": "single", "sz": "4", "space": "0", "color": COLOR_BORDER_HEX},
            left={"val": "single", "sz": "12", "space": "0", "color": accent_hex if c_idx == 0 else COLOR_BORDER_HEX},
            right={"val": "single", "sz": "4", "space": "0", "color": COLOR_BORDER_HEX},
        )
        pm = cell.paragraphs[0]
        pm.paragraph_format.space_after = Pt(2)
        r_l = pm.add_run(f"{lbl}\n")
        r_l.bold = True
        r_l.font.size = Pt(7.5)
        r_l.font.color.rgb = RGBColor(*accent_rgb)
        r_v = pm.add_run(val)
        r_v.font.size = Pt(9.5)
        r_v.font.color.rgb = RGBColor(*COLOR_DARK_TEXT_RGB)

    # ------------------------------------------------------------------
    # SECCIÓN 2: HOJAS INTERIORES LIMPIAS (CON ENCABEZADO Y PIE DE PÁGINA)
    # ------------------------------------------------------------------
    sec2 = doc.add_section(WD_SECTION.NEW_PAGE)
    sec2.page_width = Inches(8.5)
    sec2.page_height = Inches(11.0)
    sec2.top_margin = Inches(1.05)
    sec2.bottom_margin = Inches(0.9)
    sec2.left_margin = Inches(0.95)
    sec2.right_margin = Inches(0.95)
    sec2.header_distance = Inches(0.45)
    sec2.footer_distance = Inches(0.45)

    sec2.header.is_linked_to_previous = False
    sec2.footer.is_linked_to_previous = False

    # ENCABEZADO (Logo fijado por ALTURA para evitar recortes)
    header = sec2.header
    htbl = header.add_table(rows=1, cols=2, width=Inches(6.6))
    htbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    htbl.autofit = False
    remove_table_borders(htbl)

    hc_logo, hc_txt = htbl.rows[0].cells
    hc_logo.width = Inches(1.8)
    hc_txt.width = Inches(4.8)
    hc_logo.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    hc_txt.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    p_hl = hc_logo.paragraphs[0]
    p_hl.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_hl.paragraph_format.space_after = Pt(0)
    if logo_bytes:
        p_hl.add_run().add_picture(io.BytesIO(logo_bytes), height=Inches(0.42))
    else:
        r_hl = p_hl.add_run("OTEC SANTA MÓNICA")
        r_hl.bold = True
        r_hl.font.size = Pt(9)
        r_hl.font.color.rgb = RGBColor(*primary_rgb)

    p_ht = hc_txt.paragraphs[0]
    p_ht.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_ht.paragraph_format.space_after = Pt(0)
    rh1 = p_ht.add_run(f"{truncar_texto(datos['nombre_curso'], 62)}\n")
    rh1.bold = True
    rh1.font.size = Pt(8.5)
    rh1.font.color.rgb = RGBColor(*primary_rgb)
    rh2 = p_ht.add_run(f"{truncar_texto(datos['titulo_guia'], 55)}  |  OTEC Santa Mónica")
    rh2.font.size = Pt(7.5)
    rh2.font.color.rgb = RGBColor(*COLOR_MUTED_RGB)

    p_hline = header.add_paragraph()
    p_hline.paragraph_format.space_before = Pt(3)
    p_hline.paragraph_format.space_after = Pt(0)
    add_paragraph_border(p_hline, side="bottom", color_hex=accent_hex, sz="8")

    # PIE DE PÁGINA ("Página X de Y")
    footer = sec2.footer
    p_fline = footer.paragraphs[0]
    p_fline.paragraph_format.space_before = Pt(0)
    p_fline.paragraph_format.space_after = Pt(4)
    add_paragraph_border(p_fline, side="top", color_hex=COLOR_BORDER_HEX, sz="6")

    ftbl = footer.add_table(rows=1, cols=2, width=Inches(6.6))
    ftbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    ftbl.autofit = False
    remove_table_borders(ftbl)

    fc1, fc2 = ftbl.rows[0].cells
    fc1.width = Inches(5.1)
    fc2.width = Inches(1.5)

    p_f1 = fc1.paragraphs[0]
    rf1 = p_f1.add_run(f"OTEC Santa Mónica  |  {datos['pie_pagina']}")
    rf1.font.size = Pt(8)
    rf1.font.color.rgb = RGBColor(*COLOR_MUTED_RGB)

    p_f2 = fc2.paragraphs[0]
    p_f2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for txt, field in [("Página ", None), ("", "PAGE"), (" de ", None), ("", "NUMPAGES")]:
        rf = p_f2.add_run(txt)
        rf.font.size = Pt(8)
        if field:
            rf.bold = True
            rf.font.color.rgb = RGBColor(*primary_rgb)
            add_xml_field(rf, field)
        else:
            rf.font.color.rgb = RGBColor(*COLOR_MUTED_RGB)

    # Hojas interiores limpias (sin contenido)
    for i in range(datos["hojas_blancas"]):
        p_blank = doc.add_paragraph()
        p_blank.paragraph_format.space_before = Pt(0)
        p_blank.paragraph_format.space_after = Pt(0)
        if i < datos["hojas_blancas"] - 1:
            doc.add_page_break()

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ==============================================================================
# GENERADOR PDF (.PDF) — SOLO PORTADA + HOJA(S) LIMPIA(S) DE FORMATO
# ==============================================================================
class PlantillaOTECPDF(FPDF):
    def __init__(self, datos: dict, logo_bytes: bytes | None):
        super().__init__(orientation="P", unit="mm", format="Letter")
        self.datos = datos
        self.logo_bytes = logo_bytes
        self.primary_rgb = hex_to_rgb(datos["color_primario"])
        self.accent_rgb = hex_to_rgb(datos["color_acento"])
        self.alias_nb_pages()
        self.set_auto_page_break(auto=True, margin=24)
        self.set_margins(left=22, top=28, right=22)

    def header(self):
        if self.page_no() == 1:
            return

        # Logo fijado por ALTURA (10.5 mm) -> nunca se recorta ni pisa la línea en y=23
        if self.logo_bytes:
            try:
                self.image(io.BytesIO(self.logo_bytes), x=22, y=9.5, h=10.5)
            except Exception:
                pass
        else:
            self.set_xy(22, 12)
            self.set_font("Helvetica", "B", 9)
            self.set_text_color(*self.primary_rgb)
            self.cell(45, 5, "OTEC SANTA MONICA")

        # Texto derecho del encabezado
        self.set_xy(68, 10)
        self.set_font("Helvetica", "B", 8.5)
        self.set_text_color(*self.primary_rgb)
        self.cell(
            126, 4.8, limpiar_latin1(truncar_texto(self.datos["nombre_curso"], 62)),
            align="R", new_x="LMARGIN", new_y="NEXT"
        )

        self.set_x(68)
        self.set_font("Helvetica", "", 7.5)
        self.set_text_color(*COLOR_MUTED_RGB)
        sub = f"{truncar_texto(self.datos['titulo_guia'], 52)}  |  OTEC Santa Mónica"
        self.cell(126, 4.2, limpiar_latin1(sub), align="R", new_x="LMARGIN", new_y="NEXT")

        # Línea divisoria limpia
        self.set_draw_color(*self.accent_rgb)
        self.set_line_width(0.45)
        self.line(22, 23, 194, 23)
        self.set_y(28)

    def footer(self):
        if self.page_no() == 1:
            return

        self.set_y(-18)
        self.set_draw_color(*COLOR_BORDER_RGB)
        self.set_line_width(0.3)
        self.line(22, self.get_y(), 194, self.get_y())

        self.set_y(-15)
        self.set_font("Helvetica", "", 7.5)
        self.set_text_color(*COLOR_MUTED_RGB)
        pie = truncar_texto(f"OTEC Santa Mónica  |  {self.datos['pie_pagina']}", 90)
        self.cell(138, 6, limpiar_latin1(pie), align="L")

        self.set_font("Helvetica", "B", 8)
        self.set_text_color(*self.primary_rgb)
        self.cell(34, 6, limpiar_latin1(f"Página {self.page_no()} de {{nb}}"), align="R")


def generar_pdf(datos: dict, logo_bytes: bytes | None) -> bytes:
    pdf = PlantillaOTECPDF(datos, logo_bytes)
    primary_rgb = hex_to_rgb(datos["color_primario"])
    accent_rgb = hex_to_rgb(datos["color_acento"])

    # ------------------------------------------------------------------
    # PORTADA EDITORIAL ASIMÉTRICA (Página 1)
    # ------------------------------------------------------------------
    pdf.add_page()

    # Franja vertical izquierda a sangre completa (Primario + Acento)
    pdf.set_fill_color(*primary_rgb)
    pdf.rect(0, 0, 14, 279.4, style="F")
    pdf.set_fill_color(*accent_rgb)
    pdf.rect(14, 0, 2.2, 279.4, style="F")

    # Cabecera de portada: Logo (altura fija 24 mm) + Código/Fecha derecha
    if logo_bytes:
        try:
            pdf.image(io.BytesIO(logo_bytes), x=28, y=20, h=24)
        except Exception:
            pass

    pdf.set_xy(125, 22)
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.set_text_color(*primary_rgb)
    pdf.cell(69, 5, limpiar_latin1(datos["codigo"]), align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(125)
    pdf.set_font("Helvetica", "", 8.5)
    pdf.set_text_color(*COLOR_MUTED_RGB)
    pdf.cell(69, 4.5, limpiar_latin1(datos["fecha"]), align="R")

    # Bloque de título principal (alineado a la izquierda)
    pdf.set_xy(28, 86)
    pdf.set_font("Helvetica", "B", 10.5)
    pdf.set_text_color(*accent_rgb)
    pdf.cell(166, 7, limpiar_latin1(datos["tipo_documento"].upper()), align="L", new_x="LMARGIN", new_y="NEXT")

    pdf.set_x(28)
    pdf.ln(2)
    pdf.set_x(28)
    pdf.set_font("Helvetica", "B", 26)
    pdf.set_text_color(*primary_rgb)
    pdf.multi_cell(166, 11.5, limpiar_latin1(datos["nombre_curso"]), align="L")

    pdf.ln(4)
    pdf.set_x(28)
    pdf.set_font("Helvetica", "", 14)
    pdf.set_text_color(*COLOR_MUTED_RGB)
    pdf.multi_cell(166, 7.5, limpiar_latin1(datos["titulo_guia"]), align="L")

    # Línea corta de acento
    y_rule = pdf.get_y() + 6
    pdf.set_draw_color(*accent_rgb)
    pdf.set_line_width(1.2)
    pdf.line(28, y_rule, 66, y_rule)

    # Ficha técnica minimalista 2x2 en zona inferior
    placeholder = "______________________" if datos["linea_vacia"] else "Por completar"
    ficha = [
        ("RELATOR / DOCENTE", datos["relator"] or placeholder),
        ("INSTITUCIÓN / CLIENTE", datos["cliente"] or placeholder),
        ("MODALIDAD Y DURACIÓN", f"{datos['modalidad']}  -  {datos['horas']}"),
        ("ORGANISMO TÉCNICO", datos["otec_nombre"]),
    ]

    x_left = 28
    y_top = 204
    box_w = 81
    box_h = 22
    gap_x = 4
    gap_y = 4

    for idx, (lbl, val) in enumerate(ficha):
        row, col = divmod(idx, 2)
        bx = x_left + col * (box_w + gap_x)
        by = y_top + row * (box_h + gap_y)

        # Fondo suave + borde fino
        pdf.set_fill_color(*COLOR_SOFT_BG_RGB)
        pdf.set_draw_color(*COLOR_BORDER_RGB)
        pdf.set_line_width(0.25)
        pdf.rect(bx, by, box_w, box_h, style="DF")

        # Borde izquierdo de acento en cada tarjeta
        pdf.set_fill_color(*accent_rgb)
        pdf.rect(bx, by, 1.3, box_h, style="F")

        # Etiqueta
        pdf.set_xy(bx + 4, by + 3.5)
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.set_text_color(*accent_rgb)
        pdf.cell(box_w - 6, 4, limpiar_latin1(lbl))

        # Valor
        pdf.set_xy(bx + 4, by + 9.5)
        pdf.set_font("Helvetica", "", 9.2)
        pdf.set_text_color(*COLOR_DARK_TEXT_RGB)
        pdf.multi_cell(box_w - 6, 4.6, limpiar_latin1(truncar_texto(val, 48)))

    # Pie discreto de portada
    pdf.set_xy(28, 263)
    pdf.set_font("Helvetica", "", 8)
    pdf.set_text_color(*COLOR_MUTED_RGB)
    pdf.cell(166, 5, limpiar_latin1(datos["pie_pagina"]), align="L")

    # ------------------------------------------------------------------
    # HOJAS INTERIORES LIMPIAS (SIN CONTENIDO)
    # ------------------------------------------------------------------
    for _ in range(datos["hojas_blancas"]):
        pdf.add_page()

    return bytes(pdf.output())


# ==============================================================================
# INTERFAZ STREAMLIT (LISTA PARA IMPORTAR CON run())
# ==============================================================================
def run():
    st.title("Plantilla Editorial de Guías — OTEC Santa Mónica")
    st.caption(
        "Genera una plantilla limpia (Portada profesional + hojas en blanco con encabezado y pie de página) "
        "en Word (.docx) y PDF (.pdf) lista para que el relator complete sus contenidos."
    )

    logo_bytes = cargar_logo_default()

    with st.expander("Logo Institucional y Paleta de Colores", expanded=(logo_bytes is None)):
        col_logo, col_color = st.columns([1.2, 1.8])
        with col_logo:
            if logo_bytes:
                st.image(logo_bytes, width=140)
                st.caption("Logo cargado correctamente.")
            else:
                st.warning("Logo no encontrado automáticamente.")
            logo_subido = st.file_uploader("Cambiar logo (opcional)", type=["jpg", "jpeg", "png"])
            if logo_subido is not None:
                logo_bytes = logo_subido.read()
                st.image(logo_bytes, width=140)

        with col_color:
            paleta_sel = st.selectbox("Paleta de Diseño", list(PALETAS.keys()) + ["Personalizada"])
            if paleta_sel == "Personalizada":
                c1, c2 = st.columns(2)
                with c1:
                    color_primario = st.color_picker("Color Primario", "#1B365D")
                with c2:
                    color_acento = st.color_picker("Color de Acento", "#2B6CB0")
            else:
                color_primario, color_acento = PALETAS[paleta_sel]
                st.markdown(
                    f"Primario: `{color_primario}` &nbsp;|&nbsp; Acento: `{color_acento}`",
                    unsafe_allow_html=True,
                )

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
        relator = st.text_input("Nombre del Relator / Docente (opcional)", value="")
        cliente = st.text_input("Institución / Cliente (opcional)", value="")
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

    c_pie, c_hojas, c_lin = st.columns([2.4, 1, 1.2])
    with c_pie:
        pie_pagina = st.text_input(
            "Pie de Página",
            value="Material exclusivo de capacitación — Prohibida su reproducción",
        )
    with c_hojas:
        hojas_blancas = st.number_input("Hojas limpias tras portada", min_value=1, max_value=10, value=1, step=1)
    with c_lin:
        st.write("")
        st.write("")
        linea_vacia = st.checkbox("Línea (____) en vacíos", value=False)

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
        "hojas_blancas": int(hojas_blancas),
        "linea_vacia": linea_vacia,
        "color_primario": color_primario,
        "color_acento": color_acento,
    }

    st.divider()

    if st.button("Generar Plantilla (Word y PDF)", type="primary", use_container_width=True):
        if not nombre_curso.strip():
            st.error("Ingresa el nombre del curso antes de generar la plantilla.")
            return

        with st.spinner("Generando plantilla editorial en Word y PDF..."):
            docx_bytes = generar_word(datos, logo_bytes)
            pdf_bytes = generar_pdf(datos, logo_bytes)

        st.session_state["docx_plantilla"] = docx_bytes
        st.session_state["pdf_plantilla"] = pdf_bytes
        nombre_base = "".join(c if c.isalnum() or c in (" ", "-", "_") else "" for c in nombre_curso).strip()
        nombre_base = nombre_base.replace(" ", "_") or "Plantilla_OTEC_Santa_Monica"
        st.session_state["nombre_plantilla"] = f"Plantilla_{nombre_base}"
        st.success("Plantilla generada correctamente. Descárgala abajo.")

    if "docx_plantilla" in st.session_state and "pdf_plantilla" in st.session_state:
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.download_button(
                label="Descargar Plantilla Word (.docx)",
                data=st.session_state["docx_plantilla"],
                file_name=f"{st.session_state['nombre_plantilla']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with col_d2:
            st.download_button(
                label="Descargar Plantilla PDF (.pdf)",
                data=st.session_state["pdf_plantilla"],
                file_name=f"{st.session_state['nombre_plantilla']}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )


if __name__ == "__main__":
    st.set_page_config(page_title="Plantilla Editorial OTEC", layout="wide")
    run()
