#!/usr/bin/env python3
"""
Generador de la Propuesta Comercial Ejecutiva de Skooling en formato Microsoft Word (.docx).
Diseñado con precisión tipográfica, paleta institucional (Navy, Gold, White) y estructura de 4 páginas.
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --- CONSTANTES DE COLOR ---
COLOR_NAVY_DARK = RGBColor(13, 27, 42)      # #0D1B2A
COLOR_NAVY_MED  = RGBColor(30, 50, 84)      # #1E3254
COLOR_GOLD       = RGBColor(198, 161, 91)   # #C6A15B
COLOR_GOLD_DARK  = RGBColor(184, 134, 11)   # #B8860B
COLOR_TEXT_MAIN  = RGBColor(30, 41, 59)     # #1E293B
COLOR_TEXT_MUTED = RGBColor(100, 116, 139)  # #64748B
COLOR_BORDER     = RGBColor(203, 213, 224)  # #CBD5E0

HEX_NAVY_DARK = "0D1B2A"
HEX_NAVY_MED  = "1E3254"
HEX_GOLD_BG   = "FEF9E7"
HEX_GOLD_HL   = "FDE68A"
HEX_LIGHT_BG  = "F8FAFC"
HEX_BORDER    = "CBD5E0"
HEX_WHITE     = "FFFFFF"

# --- FUNCIONES DE SOPORTE XML PARA PYTHON-DOCX ---

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    """Márgenes internas de celda (en dxa: 20 dxa = 1 pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    """Define bordes de celda personalizados."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    borders = {'top': top, 'bottom': bottom, 'left': left, 'right': right}
    for b_name, b_val in borders.items():
        if b_val:
            node = OxmlElement(f'w:{b_name}')
            node.set(qn('w:val'), b_val.get('val', 'single'))
            node.set(qn('w:sz'), str(b_val.get('sz', '4')))
            node.set(qn('w:space'), '0')
            node.set(qn('w:color'), b_val.get('color', HEX_BORDER))
            tcBorders.append(node)
        else:
            node = OxmlElement(f'w:{b_name}')
            node.set(qn('w:val'), 'none')
            tcBorders.append(node)
    tcPr.append(tcBorders)

def add_page_number(run):
    """Inserta el campo PAGE en un run de texto."""
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)

def add_numpages(run):
    """Inserta el campo NUMPAGES en un run de texto."""
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "NUMPAGES"
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)

def create_section_header(doc, title_text, tag_text):
    """Crea un encabezado de sección elegante con título y tag dorado a la derecha."""
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    cell_l = tbl.cell(0, 0)
    cell_r = tbl.cell(0, 1)
    cell_l.width = Inches(5.3)
    cell_r.width = Inches(2.0)
    
    set_cell_margins(cell_l, top=60, bottom=40, left=0, right=0)
    set_cell_margins(cell_r, top=60, bottom=40, left=0, right=0)
    
    # Línea inferior de acento
    border_b = {'val': 'single', 'sz': '8', 'color': HEX_BORDER}
    border_l_accent = {'val': 'single', 'sz': '24', 'color': 'C6A15B'} # barra izquierda dorada
    
    set_cell_borders(cell_l, bottom=border_b, left=border_l_accent)
    set_cell_borders(cell_r, bottom=border_b)
    
    # Título
    p_l = cell_l.paragraphs[0]
    p_l.paragraph_format.space_before = Pt(0)
    p_l.paragraph_format.space_after = Pt(2)
    p_l.paragraph_format.left_indent = Inches(0.1)
    run_t = p_l.add_run(title_text)
    run_t.font.name = 'Arial'
    run_t.font.size = Pt(10.5)
    run_t.font.bold = True
    run_t.font.color.rgb = COLOR_NAVY_MED
    
    # Tag
    p_r = cell_r.paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_r.paragraph_format.space_before = Pt(0)
    p_r.paragraph_format.space_after = Pt(2)
    run_tag = p_r.add_run(tag_text)
    run_tag.font.name = 'Arial'
    run_tag.font.size = Pt(7.5)
    run_tag.font.bold = True
    run_tag.font.color.rgb = COLOR_GOLD_DARK

def format_cell_text(cell, bold_prefix="", text="", size=8.5, color=COLOR_TEXT_MAIN, space_after=2, new_paragraph=False):
    if new_paragraph or len(cell.paragraphs) == 0:
        p = cell.add_paragraph()
    else:
        p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_b = p.add_run(bold_prefix)
        r_b.font.name = 'Arial'
        r_b.font.size = Pt(size)
        r_b.font.bold = True
        r_b.font.color.rgb = color
    if text:
        r_t = p.add_run(text)
        r_t.font.name = 'Arial'
        r_t.font.size = Pt(size)
        r_t.font.color.rgb = color
    return p

def add_bullet_item(cell, title="", text="", size=8.0, check_color=COLOR_GOLD_DARK):
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.1
    p.paragraph_format.left_indent = Inches(0.18)
    p.paragraph_format.first_line_indent = Inches(-0.18)
    
    # Checkmark icon
    r_ck = p.add_run("✔ ")
    r_ck.font.name = 'Arial'
    r_ck.font.size = Pt(size)
    r_ck.font.bold = True
    r_ck.font.color.rgb = check_color
    
    if title:
        r_t = p.add_run(title + " ")
        r_t.font.name = 'Arial'
        r_t.font.size = Pt(size)
        r_t.font.bold = True
        r_t.font.color.rgb = COLOR_TEXT_MAIN
        
    if text:
        r_d = p.add_run(text)
        r_d.font.name = 'Arial'
        r_d.font.size = Pt(size)
        r_d.font.color.rgb = COLOR_TEXT_MAIN

# ==============================================================================
# CONSTRUCTOR PRINCIPAL DEL DOCUMENTO DOCX
# ==============================================================================

def generate_commercial_docx(output_path, assets_dir):
    doc = docx.Document()
    
    # 1. Configuración de Página (Letter, Márgenes de 0.6 in)
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.6)
    section.right_margin = Inches(0.6)
    section.header_distance = Inches(0.3)
    section.footer_distance = Inches(0.3)
    
    # 2. Configuración del Encabezado (Header con Logos y Metadatos)
    header = section.header
    header_tbl = header.add_table(rows=1, cols=2, width=Inches(7.3))
    header_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_tbl.autofit = False
    
    h_cell_l = header_tbl.cell(0, 0)
    h_cell_r = header_tbl.cell(0, 1)
    h_cell_l.width = Inches(4.0)
    h_cell_r.width = Inches(3.3)
    
    set_cell_margins(h_cell_l, top=0, bottom=40, left=0, right=0)
    set_cell_margins(h_cell_r, top=0, bottom=40, left=0, right=0)
    
    border_bottom_header = {'val': 'single', 'sz': '6', 'color': HEX_BORDER}
    set_cell_borders(h_cell_l, bottom=border_bottom_header)
    set_cell_borders(h_cell_r, bottom=border_bottom_header)
    
    p_hl = h_cell_l.paragraphs[0]
    p_hl.paragraph_format.space_before = Pt(0)
    p_hl.paragraph_format.space_after = Pt(2)
    
    # Agregar logo Aura
    aura_logo_path = os.path.join(assets_dir, "Mountain_3.png")
    skooling_logo_path = os.path.join(assets_dir, "logo-full-Route.png")
    
    if os.path.exists(aura_logo_path):
        p_hl.add_run().add_picture(aura_logo_path, height=Inches(0.40))
        r_sep = p_hl.add_run("   |   ")
        r_sep.font.name = 'Arial'
        r_sep.font.size = Pt(14)
        r_sep.font.color.rgb = COLOR_BORDER
    
    if os.path.exists(skooling_logo_path):
        p_hl.add_run().add_picture(skooling_logo_path, height=Inches(0.38))
        
    p_hr = h_cell_r.paragraphs[0]
    p_hr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_hr.paragraph_format.space_before = Pt(0)
    p_hr.paragraph_format.space_after = Pt(2)
    p_hr.paragraph_format.line_spacing = 1.15
    
    r_meta1 = p_hr.add_run("Propuesta Comercial Ejecutiva\n")
    r_meta1.font.name = 'Arial'
    r_meta1.font.size = Pt(8.5)
    r_meta1.font.bold = True
    r_meta1.font.color.rgb = COLOR_NAVY_MED
    
    r_meta2 = p_hr.add_run("Dirigido a: Rectoría, Junta Directiva y Dirección Administrativa\n")
    r_meta2.font.name = 'Arial'
    r_meta2.font.size = Pt(7.5)
    r_meta2.font.color.rgb = COLOR_TEXT_MUTED
    
    r_meta3 = p_hr.add_run("Respaldado por: Aura Soluciones e Innovación S.A.S. (aurasi.co)")
    r_meta3.font.name = 'Arial'
    r_meta3.font.size = Pt(7.5)
    r_meta3.font.bold = True
    r_meta3.font.color.rgb = COLOR_GOLD_DARK

    # 3. Configuración del Pie de Página (Footer con Paginación y Contacto)
    footer = section.footer
    f_tbl = footer.add_table(rows=1, cols=3, width=Inches(7.3))
    f_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    f_tbl.autofit = False
    
    f_c1 = f_tbl.cell(0, 0)
    f_c2 = f_tbl.cell(0, 1)
    f_c3 = f_tbl.cell(0, 2)
    f_c1.width = Inches(3.2)
    f_c2.width = Inches(2.3)
    f_c3.width = Inches(1.8)
    
    set_cell_margins(f_c1, top=40, bottom=0, left=0, right=0)
    set_cell_margins(f_c2, top=40, bottom=0, left=0, right=0)
    set_cell_margins(f_c3, top=40, bottom=0, left=0, right=0)
    
    border_top_footer = {'val': 'single', 'sz': '4', 'color': HEX_BORDER}
    set_cell_borders(f_c1, top=border_top_footer)
    set_cell_borders(f_c2, top=border_top_footer)
    set_cell_borders(f_c3, top=border_top_footer)
    
    p_f1 = f_c1.paragraphs[0]
    p_f1.paragraph_format.space_before = Pt(2)
    r_f1 = p_f1.add_run("Aura Soluciones e Innovación S.A.S. • Skooling")
    r_f1.font.name = 'Arial'
    r_f1.font.size = Pt(7.5)
    r_f1.font.bold = True
    r_f1.font.color.rgb = COLOR_NAVY_MED
    
    p_f2 = f_c2.paragraphs[0]
    p_f2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_f2.paragraph_format.space_before = Pt(2)
    r_f2 = p_f2.add_run("aurasi.co | contacto@aurasi.co")
    r_f2.font.name = 'Arial'
    r_f2.font.size = Pt(7.5)
    r_f2.font.color.rgb = COLOR_TEXT_MUTED
    
    p_f3 = f_c3.paragraphs[0]
    p_f3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_f3.paragraph_format.space_before = Pt(2)
    r_f3 = p_f3.add_run("Página ")
    r_f3.font.name = 'Arial'
    r_f3.font.size = Pt(7.5)
    r_f3.font.color.rgb = COLOR_NAVY_MED
    add_page_number(r_f3)
    r_f3b = p_f3.add_run(" de ")
    r_f3b.font.name = 'Arial'
    r_f3b.font.size = Pt(7.5)
    r_f3b.font.color.rgb = COLOR_NAVY_MED
    add_numpages(r_f3b)

    # ==========================================================================
    # PÁGINA 1: CASO DE NEGOCIO Y DIAGNÓSTICO COMPARATIVO
    # ==========================================================================
    
    # Banner Principal Hero
    banner_tbl = doc.add_table(rows=1, cols=1)
    banner_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    banner_tbl.autofit = False
    b_cell = banner_tbl.cell(0, 0)
    b_cell.width = Inches(7.3)
    set_cell_background(b_cell, HEX_LIGHT_BG)
    set_cell_margins(b_cell, top=140, bottom=140, left=180, right=180)
    set_cell_borders(b_cell, 
                     top={'val': 'single', 'sz': '4', 'color': HEX_BORDER},
                     bottom={'val': 'single', 'sz': '4', 'color': HEX_BORDER},
                     left={'val': 'single', 'sz': '24', 'color': 'C6A15B'},
                     right={'val': 'single', 'sz': '4', 'color': HEX_BORDER})
    
    p_b0 = b_cell.paragraphs[0]
    p_b0.paragraph_format.space_before = Pt(0)
    p_b0.paragraph_format.space_after = Pt(2)
    r_sub = p_b0.add_run("ECOSISTEMA EDUCATIVO SAAS TODO-EN-UNO")
    r_sub.font.name = 'Arial'
    r_sub.font.size = Pt(8.0)
    r_sub.font.bold = True
    r_sub.font.color.rgb = COLOR_GOLD_DARK
    
    p_b1 = b_cell.add_paragraph()
    p_b1.paragraph_format.space_before = Pt(1)
    p_b1.paragraph_format.space_after = Pt(4)
    r_title = p_b1.add_run("SKOOLING: La vida escolar de su comunidad, centralizada en un solo lugar")
    r_title.font.name = 'Arial'
    r_title.font.size = Pt(13.0)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_NAVY_DARK
    
    p_b2 = b_cell.add_paragraph()
    p_b2.paragraph_format.space_before = Pt(0)
    p_b2.paragraph_format.space_after = Pt(0)
    p_b2.paragraph_format.line_spacing = 1.15
    r_desc1 = p_b2.add_run("Plataforma integral en la nube que unifica la ")
    r_desc1.font.name = 'Arial'
    r_desc1.font.size = Pt(8.5)
    r_desc1.font.color.rgb = COLOR_TEXT_MAIN
    
    r_bld1 = p_b2.add_run("gestión académica")
    r_bld1.font.name = 'Arial'
    r_bld1.font.size = Pt(8.5)
    r_bld1.font.bold = True
    r_bld1.font.color.rgb = COLOR_TEXT_MAIN
    
    r_desc2 = p_b2.add_run(", la ")
    r_desc2.font.name = 'Arial'
    r_desc2.font.size = Pt(8.5)
    r_desc2.font.color.rgb = COLOR_TEXT_MAIN
    
    r_bld2 = p_b2.add_run("comunicación institucional certificada")
    r_bld2.font.name = 'Arial'
    r_bld2.font.size = Pt(8.5)
    r_bld2.font.bold = True
    r_bld2.font.color.rgb = COLOR_TEXT_MAIN
    
    r_desc3 = p_b2.add_run(" y el ")
    r_desc3.font.name = 'Arial'
    r_desc3.font.size = Pt(8.5)
    r_desc3.font.color.rgb = COLOR_TEXT_MAIN
    
    r_bld3 = p_b2.add_run("monitoreo satelital de rutas escolares (GPS)")
    r_bld3.font.name = 'Arial'
    r_bld3.font.size = Pt(8.5)
    r_bld3.font.bold = True
    r_bld3.font.color.rgb = COLOR_TEXT_MAIN
    
    r_desc4 = p_b2.add_run(" en un único Portal Web Directivo/Docente y una única App Móvil para las familias.")
    r_desc4.font.name = 'Arial'
    r_desc4.font.size = Pt(8.5)
    r_desc4.font.color.rgb = COLOR_TEXT_MAIN

    # Sección 1: Dos Realidades
    create_section_header(doc, "1. Caso de Negocio para el Decisor: Dos Realidades Institucionales", "VALOR Y BENEFICIO INSTITUCIONAL")
    
    grid1_tbl = doc.add_table(rows=1, cols=2)
    grid1_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    grid1_tbl.autofit = False
    c_g1 = grid1_tbl.cell(0, 0)
    c_g2 = grid1_tbl.cell(0, 1)
    c_g1.width = Inches(3.55)
    c_g2.width = Inches(3.55)
    
    # Cuadro A: Sin Plataforma
    set_cell_background(c_g1, HEX_LIGHT_BG)
    set_cell_margins(c_g1, top=100, bottom=100, left=120, right=120)
    set_cell_borders(c_g1, 
                     top={'val': 'single', 'sz': '4', 'color': HEX_BORDER},
                     bottom={'val': 'single', 'sz': '4', 'color': HEX_BORDER},
                     left={'val': 'single', 'sz': '16', 'color': 'B91C1C'}, # Borde rojo sutil
                     right={'val': 'single', 'sz': '4', 'color': HEX_BORDER})
    
    format_cell_text(c_g1, bold_prefix="A. Si su colegio aún no tiene plataforma digital", color=RGBColor(185, 28, 28), size=9.0)
    p_ca = c_g1.add_paragraph()
    p_ca.paragraph_format.space_before = Pt(2)
    p_ca.paragraph_format.space_after = Pt(4)
    p_ca.paragraph_format.line_spacing = 1.15
    r_ca = p_ca.add_run("Para una institución que opera con Excel, circulares físicas y grupos de WhatsApp, implementar Skooling ")
    r_ca.font.name = 'Arial'; r_ca.font.size = Pt(8.0)
    r_cab = p_ca.add_run("no es un gasto adicional, sino una inversión con valor directo:")
    r_cab.font.name = 'Arial'; r_cab.font.size = Pt(8.0); r_cab.font.bold = True
    
    add_bullet_item(c_g1, "Valor agregado para la comunidad:", "La comunicación directa y oportuna entre la institución y las familias en todos los aspectos se reconoce como un sello de calidad para el colegio.", size=7.8)
    add_bullet_item(c_g1, "Ahorro operativo y cero papel:", "Automatiza boletines en PDF, horarios, control de asistencia y comunicados en segundos.", size=7.8)
    add_bullet_item(c_g1, "Protección laboral docente:", "Elimina el uso del WhatsApp personal de los profesores fuera del horario laboral.", size=7.8)
    add_bullet_item(c_g1, "Adopción progresiva:", "Permite iniciar desde el Nivel 1 (Core Académico) con una inversión anual ajustada.", size=7.8)

    # Cuadro B: Plataformas Fragmentadas
    set_cell_background(c_g2, HEX_GOLD_BG)
    set_cell_margins(c_g2, top=100, bottom=100, left=120, right=120)
    set_cell_borders(c_g2, 
                     top={'val': 'single', 'sz': '4', 'color': 'D4AF37'},
                     bottom={'val': 'single', 'sz': '4', 'color': 'D4AF37'},
                     left={'val': 'single', 'sz': '20', 'color': 'C6A15B'}, # Borde dorado
                     right={'val': 'single', 'sz': '4', 'color': 'D4AF37'})
    
    format_cell_text(c_g2, bold_prefix="B. Si su colegio hoy usa plataformas fragmentadas", color=COLOR_NAVY_DARK, size=9.0)
    p_cb = c_g2.add_paragraph()
    p_cb.paragraph_format.space_before = Pt(2)
    p_cb.paragraph_format.space_after = Pt(4)
    p_cb.paragraph_format.line_spacing = 1.15
    r_cb = p_cb.add_run("Muchos colegios sufren hoy la dispersión tecnológica: publican avisos en su web, usan otra para notas, otra para rutas y aun así usan WhatsApp:")
    r_cb.font.name = 'Arial'; r_cb.font.size = Pt(8.0)
    
    add_bullet_item(c_g2, "Fin a la información 'regada':", "Cuando un padre debe revisar 3 o 4 fuentes distintas, termina perdiendo comunicados. Skooling crea una Única Fuente de la Verdad.", size=7.8)
    add_bullet_item(c_g2, "1 Sola App para Académico + Rutas:", "La plataforma del colegio por fin integra en vivo la ubicación satelital GPS del bus escolar.", size=7.8)
    add_bullet_item(c_g2, "Economía en un único contrato:", "Unificar la operación en una sola plataforma evita pagar licencias y cobros separados a múltiples proveedores.", size=7.8)

    # Sección 2: Tabla Comparativa
    create_section_header(doc, "2. Comparativa de Impacto Operativo, Comunicacional y Financiero", "DIAGNÓSTICO COMPARATIVO")
    
    comp_tbl = doc.add_table(rows=6, cols=4)
    comp_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    comp_tbl.autofit = False
    
    col_widths = [Inches(1.5), Inches(1.85), Inches(1.85), Inches(2.1)]
    for row in comp_tbl.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = width
            set_cell_margins(row.cells[idx], top=60, bottom=60, left=80, right=80)
            
    # Headers
    headers_text = [
        "Criterio Evaluado",
        "Operación Manual / Tradicional",
        "Sistemas Fragmentados (Múltiples Apps)",
        "Ecosistema Centralizado SKOOLING"
    ]
    for idx, text in enumerate(headers_text):
        cell = comp_tbl.cell(0, idx)
        if idx == 3:
            set_cell_background(cell, HEX_NAVY_MED)
            set_cell_borders(cell, top={'color': 'C6A15B', 'sz': '8'}, bottom={'color': 'C6A15B', 'sz': '8'})
            format_cell_text(cell, bold_prefix=text, color=RGBColor(255, 215, 0), size=7.8)
        else:
            set_cell_background(cell, HEX_NAVY_DARK)
            set_cell_borders(cell, top={'color': HEX_BORDER, 'sz': '4'}, bottom={'color': HEX_BORDER, 'sz': '4'})
            format_cell_text(cell, bold_prefix=text, color=RGBColor(255, 255, 255), size=7.8)

    # Datos Filas
    rows_data = [
        ("Fuente de Información para Padres",
         "Agenda de papel, llamadas a secretaría y rumores en grupos de WhatsApp.",
         "Información 'regada' entre página web, plataforma de notas, app de transporte y WhatsApp.",
         "1 Sola App oficial con todo: Clases, Agenda, Boletines, Mensajes y Rutas."),
        ("Monitoreo de Rutas Escolares",
         "Incertidumbre total; llamadas constantes a portería por retrasos.",
         "Desconectado del sistema del colegio; requiere contratar otra aplicación externa por aparte.",
         "Integrado nativamente con mapa GPS en vivo y alertas de cercanía."),
        ("Efectividad en Comunicaciones",
         "Circulares impresas extraviadas; sin certeza de quién leyó el aviso.",
         "Los padres omiten anuncios por fatiga digital al tener múltiples canales y claves.",
         "Broadcasts oficiales con notificación Push y confirmación de lectura."),
        ("Privacidad y Orden Institucional",
         "Teléfonos personales de docentes expuestos a mensajes a deshoras.",
         "Experiencia inconsistente; a pesar de pagar software se sigue usando WhatsApp.",
         "Canal institucional 100% blindado, trazable y de alto prestigio."),
        ("Costo Total Anual",
         "Alto costo oculto en papelería, reprocesos y desgaste administrativo.",
         "Sobrecosto elevado: Doble o triple cobro (Licencia académica + App de rutas aparte).",
         "1 Único Contrato Anual centralizado, más económico que pagar proveedores separados.")
    ]

    border_row = {'val': 'single', 'sz': '4', 'color': HEX_BORDER}
    for row_idx, data in enumerate(rows_data, start=1):
        is_highlight = (row_idx == 5)
        for col_idx, text in enumerate(data):
            cell = comp_tbl.cell(row_idx, col_idx)
            set_cell_borders(cell, top=border_row, bottom=border_row, left=border_row, right=border_row)
            
            if col_idx == 0:
                format_cell_text(cell, bold_prefix=text, color=COLOR_TEXT_MAIN, size=7.5)
            elif col_idx == 3:
                bg = HEX_GOLD_HL if is_highlight else HEX_GOLD_BG
                set_cell_background(cell, bg)
                format_cell_text(cell, bold_prefix=text if is_highlight else "", text="" if is_highlight else text,
                                 color=COLOR_NAVY_DARK if is_highlight else COLOR_TEXT_MAIN, size=7.5)
            else:
                set_cell_background(cell, HEX_WHITE)
                format_cell_text(cell, text=text, color=COLOR_TEXT_MAIN, size=7.5)

    # ==========================================================================
    # PÁGINA 2: ARQUITECTURA MODULAR, RECORRIDO 360 Y EVIDENCIA VISUAL
    # ==========================================================================
    doc.add_page_break()

    # Sección 3: Arquitectura Modular Progresiva
    create_section_header(doc, "3. Arquitectura Modular Progresiva: Tres Niveles Adaptados al Colegio", "ALCANCE POR MÓDULO")
    
    p_arch_intro = doc.add_paragraph()
    p_arch_intro.paragraph_format.space_before = Pt(2)
    p_arch_intro.paragraph_format.space_after = Pt(4)
    r_ai = p_arch_intro.add_run("La institución puede contratar desde la base operativa académica hasta el ecosistema completo todo-en-uno. Cada nivel superior incluye todas las funcionalidades del nivel anterior:")
    r_ai.font.name = 'Arial'; r_ai.font.size = Pt(8.0); r_ai.font.color.rgb = COLOR_TEXT_MUTED
    
    grid3_tbl = doc.add_table(rows=1, cols=3)
    grid3_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    grid3_tbl.autofit = False
    
    c_n1 = grid3_tbl.cell(0, 0)
    c_n2 = grid3_tbl.cell(0, 1)
    c_n3 = grid3_tbl.cell(0, 2)
    c_n1.width = Inches(2.4)
    c_n2.width = Inches(2.4)
    c_n3.width = Inches(2.5)
    
    # Nivel 1
    set_cell_background(c_n1, HEX_LIGHT_BG)
    set_cell_margins(c_n1, top=80, bottom=80, left=90, right=90)
    set_cell_borders(c_n1, top={'color': HEX_BORDER, 'sz': '4'}, bottom={'color': HEX_BORDER, 'sz': '4'},
                           left={'color': '1E3254', 'sz': '12'}, right={'color': HEX_BORDER, 'sz': '4'})
    
    format_cell_text(c_n1, bold_prefix="NIVEL 1 • BASE DIGITAL\n", text="Core Académico", color=COLOR_NAVY_MED, size=8.5)
    add_bullet_item(c_n1, text="Gestión de Cursos y Asignaturas.", size=7.5)
    add_bullet_item(c_n1, text="Sistema de calificación configurable (escalas y ponderaciones).", size=7.5)
    add_bullet_item(c_n1, text="Generación masiva de boletines oficiales en PDF.", size=7.5)
    add_bullet_item(c_n1, text="Horarios semanales con validación de cruces.", size=7.5)
    add_bullet_item(c_n1, text="Control de asistencia y faltas justificadas.", size=7.5)
    add_bullet_item(c_n1, text="Cargue masivo en 1 clic de alumnos y docentes.", size=7.5)

    # Nivel 2
    set_cell_background(c_n2, HEX_GOLD_BG)
    set_cell_margins(c_n2, top=80, bottom=80, left=90, right=90)
    set_cell_borders(c_n2, top={'color': 'C6A15B', 'sz': '4'}, bottom={'color': 'C6A15B', 'sz': '4'},
                           left={'color': 'C6A15B', 'sz': '16'}, right={'color': 'C6A15B', 'sz': '4'})
    
    format_cell_text(c_n2, bold_prefix="NIVEL 2 • CERO WHATSAPP\n", text="Core + Comunicación Oficial", color=COLOR_GOLD_DARK, size=8.5)
    add_bullet_item(c_n2, text="Incluye 100% del Core Académico.", size=7.5)
    add_bullet_item(c_n2, text="Circulares digitales (Broadcasts) masivas o segmentadas.", size=7.5)
    add_bullet_item(c_n2, text="Confirmación de lectura certificada en tiempo real.", size=7.5)
    add_bullet_item(c_n2, text="Mensajería oficial profesor-padre sin números personales.", size=7.5)
    add_bullet_item(c_n2, text="Calendario Maestro Web de eventos y tareas.", size=7.5)
    add_bullet_item(c_n2, text="Notificaciones Push nativas e ingreso seguro a 1 clic.", size=7.5)

    # Nivel 3
    set_cell_background(c_n3, HEX_NAVY_DARK)
    set_cell_margins(c_n3, top=80, bottom=80, left=90, right=90)
    set_cell_borders(c_n3, top={'color': 'FFD700', 'sz': '6'}, bottom={'color': 'FFD700', 'sz': '6'},
                           left={'color': 'FFD700', 'sz': '16'}, right={'color': 'FFD700', 'sz': '6'})
    
    format_cell_text(c_n3, bold_prefix="NIVEL 3 • UNIFICACIÓN TOTAL ★\n", text="Full: Core + Mensajería + Rutas", color=RGBColor(255, 215, 0), size=8.5)
    add_bullet_item(c_n3, text="Incluye Todo Core + Mensajería.", size=7.5, check_color=RGBColor(255, 215, 0))
    add_bullet_item(c_n3, text="Skooling RouteManager: Mapa satelital en vivo GPS del bus.", size=7.5, check_color=RGBColor(255, 215, 0))
    add_bullet_item(c_n3, text="Alertas automáticas de proximidad al paradero.", size=7.5, check_color=RGBColor(255, 215, 0))
    add_bullet_item(c_n3, text="App para conductores y monitores con lista de abordaje.", size=7.5, check_color=RGBColor(255, 215, 0))
    add_bullet_item(c_n3, text="Elimina pagos a apps externas y llamadas a portería.", size=7.5, check_color=RGBColor(255, 215, 0))

    # Sección 4: Experiencia 360 en 5 Etapas
    create_section_header(doc, "4. Experiencia 360°: Cómo Skooling Acompaña Cada Momento de la Jornada", "1 SOLA APP PARA EL PADRE")
    
    step_tbl = doc.add_table(rows=1, cols=5)
    step_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    step_tbl.autofit = False
    
    steps_data = [
        ("01. La Mañana", "El padre consulta en la App el horario, uniforme, tareas y materiales del día."),
        ("02. El Trayecto", "Observa en el mapa GPS por dónde viene la ruta y recibe alerta de llegada."),
        ("03. En Colegio", "Docentes registran asistencia, notas parciales y observador digital."),
        ("04. Comunicación", "Rectoría emite circulares oficiales con acuse de lectura y calendario."),
        ("05. El Retorno", "Notificación Push al abordar la ruta de regreso y monitoreo satelital en vivo.")
    ]
    
    step_width = Inches(7.3 / 5.0)
    for idx, (stitle, sdesc) in enumerate(steps_data):
        cell = step_tbl.cell(0, idx)
        cell.width = step_width
        set_cell_background(cell, HEX_LIGHT_BG)
        set_cell_margins(cell, top=60, bottom=60, left=70, right=70)
        set_cell_borders(cell, top={'color': 'C6A15B', 'sz': '8'}, bottom={'color': HEX_BORDER, 'sz': '4'},
                               left={'color': HEX_BORDER, 'sz': '4'}, right={'color': HEX_BORDER, 'sz': '4'})
        
        format_cell_text(cell, bold_prefix=stitle + "\n", color=COLOR_NAVY_MED, size=8.0)
        p_st = cell.paragraphs[0]
        r_st = p_st.add_run(sdesc)
        r_st.font.name = 'Arial'; r_st.font.size = Pt(7.0); r_st.font.color.rgb = COLOR_TEXT_MUTED

    # Sección 5: Evidencia Visual de Interfaces
    create_section_header(doc, "5. Interfaz Moderna e Intuitiva: Portal Web Directivo + App Móvil Familias", "INTERFACES REALES")
    
    img_tbl = doc.add_table(rows=1, cols=3)
    img_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    img_tbl.autofit = False
    
    c_img1 = img_tbl.cell(0, 0)
    c_img2 = img_tbl.cell(0, 1)
    c_img3 = img_tbl.cell(0, 2)
    c_img1.width = Inches(3.3)
    c_img2.width = Inches(2.0)
    c_img3.width = Inches(2.0)
    
    for cell in [c_img1, c_img2, c_img3]:
        set_cell_background(cell, HEX_LIGHT_BG)
        set_cell_margins(cell, top=60, bottom=60, left=60, right=60)
        set_cell_borders(cell, top={'color': HEX_BORDER, 'sz': '4'}, bottom={'color': HEX_BORDER, 'sz': '4'},
                               left={'color': HEX_BORDER, 'sz': '4'}, right={'color': HEX_BORDER, 'sz': '4'})
    
    # Imagen 1: Web Calendar
    cal_img_path = os.path.join(assets_dir, "calendar.png")
    if os.path.exists(cal_img_path):
        p_i1 = c_img1.paragraphs[0]
        p_i1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_i1.paragraph_format.space_before = Pt(0)
        p_i1.paragraph_format.space_after = Pt(2)
        p_i1.add_run().add_picture(cal_img_path, width=Inches(3.1))
        p_cap1 = format_cell_text(c_img1, bold_prefix="Portal Web: ", text="Calendario Maestro y Gestión Académica", size=7.2, color=COLOR_NAVY_MED, new_paragraph=True)
        p_cap1.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Imagen 2: App Mobile
    mob_img_path = os.path.join(assets_dir, "Mobile_Sample.png")
    if os.path.exists(mob_img_path):
        p_i2 = c_img2.paragraphs[0]
        p_i2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_i2.paragraph_format.space_before = Pt(0)
        p_i2.paragraph_format.space_after = Pt(2)
        p_i2.add_run().add_picture(mob_img_path, height=Inches(1.8))
        p_cap2 = format_cell_text(c_img2, bold_prefix="App Familias: ", text="5 Módulos Unificados", size=7.2, color=COLOR_NAVY_MED, new_paragraph=True)
        p_cap2.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Imagen 3: Rutas GPS
    routes_img_path = os.path.join(assets_dir, "Routes.png")
    if os.path.exists(routes_img_path):
        p_i3 = c_img3.paragraphs[0]
        p_i3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_i3.paragraph_format.space_before = Pt(0)
        p_i3.paragraph_format.space_after = Pt(2)
        p_i3.add_run().add_picture(routes_img_path, height=Inches(1.8))
        p_cap3 = format_cell_text(c_img3, bold_prefix="Rutas Escolares: ", text="Telemetría GPS Satelital", size=7.2, color=COLOR_NAVY_MED, new_paragraph=True)
        p_cap3.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ==========================================================================
    # PÁGINA 3: PROPUESTA ECONÓMICA Y ANÁLISIS DE EFICIENCIA
    # ==========================================================================
    doc.add_page_break()

    # Sección 6: Matriz de Precios
    create_section_header(doc, "6. Tabla Institucional de Inversión Anual por Rango de Estudiantes", "TARIFA PLANA ANUAL")
    
    p_pr_intro = doc.add_paragraph()
    p_pr_intro.paragraph_format.space_before = Pt(2)
    p_pr_intro.paragraph_format.space_after = Pt(4)
    r_pri = p_pr_intro.add_run("Nuestro modelo de licenciamiento institucional ofrece una tarifa anual cerrada y predecible según el rango de estudiantes, sin cobros sorpresivos por alumno adicional dentro del tramo contratado:")
    r_pri.font.name = 'Arial'; r_pri.font.size = Pt(8.0); r_pri.font.color.rgb = COLOR_TEXT_MUTED
    
    price_tbl = doc.add_table(rows=5, cols=4)
    price_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    price_tbl.autofit = False
    
    p_col_widths = [Inches(2.1), Inches(1.65), Inches(1.75), Inches(1.8)]
    for row in price_tbl.rows:
        for idx, width in enumerate(p_col_widths):
            row.cells[idx].width = width
            set_cell_margins(row.cells[idx], top=70, bottom=70, left=80, right=80)
            
    p_headers = [
        "Rango de Estudiantes",
        "Plan 1: Core Académico\n(Cursos, Notas, Boletines, Horarios)",
        "Plan 2: Core + Mensajería\n(+ Broadcasts, Acuse, Agenda)",
        "Plan 3: Full Todo-en-Uno ★\n(+ Rutas GPS en Vivo y Monitores)"
    ]
    for idx, text in enumerate(p_headers):
        cell = price_tbl.cell(0, idx)
        if idx == 3:
            set_cell_background(cell, HEX_NAVY_MED)
            set_cell_borders(cell, top={'color': 'C6A15B', 'sz': '8'}, bottom={'color': 'C6A15B', 'sz': '8'})
            format_cell_text(cell, bold_prefix=text, color=RGBColor(255, 215, 0), size=7.5)
        else:
            set_cell_background(cell, HEX_NAVY_DARK)
            set_cell_borders(cell, top={'color': HEX_BORDER, 'sz': '4'}, bottom={'color': HEX_BORDER, 'sz': '4'})
            format_cell_text(cell, bold_prefix=text, color=RGBColor(255, 255, 255), size=7.5)

    p_rows = [
        ("Rango 1: Hasta 150 Estudiantes", "$4'800.000 COP / año", "$7'200.000 COP / año", "$9'800.000 COP / año"),
        ("Rango 2: 151 a 250 Estudiantes (Estándar)", "$6'500.000 COP / año", "$9'500.000 COP / año", "$12'500.000 COP / año"),
        ("Rango 3: 251 a 400 Estudiantes", "$11'000.000 COP / año", "$16'000.000 COP / año", "$22'000.000 COP / año"),
        ("Rango 4: Más de 400 Estudiantes", "Cotización corporativa a la medida con tarifa preferencial por volumen de estudiantes.", "", "")
    ]

    for r_idx, row_data in enumerate(p_rows, start=1):
        is_r2 = (r_idx == 2)
        is_r4 = (r_idx == 4)
        
        if is_r4:
            # Merge de celdas 1, 2, 3
            c0 = price_tbl.cell(r_idx, 0)
            c1 = price_tbl.cell(r_idx, 1)
            c2 = price_tbl.cell(r_idx, 2)
            c3 = price_tbl.cell(r_idx, 3)
            c1.merge(c3)
            
            set_cell_borders(c0, top=border_row, bottom=border_row, left=border_row, right=border_row)
            set_cell_borders(c1, top=border_row, bottom=border_row, left=border_row, right=border_row)
            format_cell_text(c0, bold_prefix=row_data[0], color=COLOR_TEXT_MAIN, size=8.0)
            format_cell_text(c1, bold_prefix=row_data[1], color=COLOR_NAVY_MED, size=8.0)
            c1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            continue
            
        for c_idx, val in enumerate(row_data):
            cell = price_tbl.cell(r_idx, c_idx)
            set_cell_borders(cell, top=border_row, bottom=border_row, left=border_row, right=border_row)
            
            if is_r2:
                bg = HEX_GOLD_HL if c_idx == 3 else HEX_GOLD_BG
                set_cell_background(cell, bg)
                format_cell_text(cell, bold_prefix=val, color=COLOR_NAVY_DARK, size=8.2)
            else:
                if c_idx == 3:
                    set_cell_background(cell, HEX_GOLD_BG)
                    format_cell_text(cell, bold_prefix=val, color=COLOR_NAVY_DARK, size=8.0)
                else:
                    set_cell_background(cell, HEX_WHITE)
                    format_cell_text(cell, bold_prefix=val if c_idx > 0 else "", text="" if c_idx > 0 else val, color=COLOR_TEXT_MAIN, size=8.0)

    # Sección 7: Análisis Financiero y Beneficios
    create_section_header(doc, "7. Análisis de Eficiencia Económica para el Comité Directivo", "POR QUÉ ES MÁS RENTABLE")
    
    an_tbl = doc.add_table(rows=1, cols=2)
    an_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    an_tbl.autofit = False
    
    c_an1 = an_tbl.cell(0, 0)
    c_an2 = an_tbl.cell(0, 1)
    c_an1.width = Inches(3.55)
    c_an2.width = Inches(3.55)
    
    # Cuadro 1: Ahorro por Unificación
    set_cell_background(c_an1, HEX_GOLD_BG)
    set_cell_margins(c_an1, top=90, bottom=90, left=110, right=110)
    set_cell_borders(c_an1, top={'color': 'C6A15B', 'sz': '4'}, bottom={'color': 'C6A15B', 'sz': '4'},
                            left={'color': 'C6A15B', 'sz': '16'}, right={'color': 'C6A15B', 'sz': '4'})
    
    format_cell_text(c_an1, bold_prefix="💰 Economía por unificación en un único contrato", color=COLOR_NAVY_DARK, size=8.8)
    p_an1 = c_an1.add_paragraph()
    p_an1.paragraph_format.space_before = Pt(2)
    p_an1.paragraph_format.space_after = Pt(4)
    p_an1.paragraph_format.line_spacing = 1.15
    r_an1 = p_an1.add_run("En colegios que contratan por un lado la plataforma académica y por otro una app de transporte, el costo combinado supera con creces nuestra tarifa unificada:")
    r_an1.font.name = 'Arial'; r_an1.font.size = Pt(7.8)
    
    add_bullet_item(c_an1, "Un solo contrato, toda la información centralizada:", 
                    "La unificación de la plataforma brinda una economía directa para el colegio al asumir un único convenio anual que integra gestión académica, mensajería oficial y monitoreo satelital GPS, eliminando cobros separados a múltiples proveedores.", size=7.6)

    # Cuadro 2: Valor Agregado Sin Plataforma
    set_cell_background(c_an2, HEX_LIGHT_BG)
    set_cell_margins(c_an2, top=90, bottom=90, left=110, right=110)
    set_cell_borders(c_an2, top={'color': HEX_BORDER, 'sz': '4'}, bottom={'color': HEX_BORDER, 'sz': '4'},
                            left={'color': '1E3254', 'sz': '16'}, right={'color': HEX_BORDER, 'sz': '4'})
    
    format_cell_text(c_an2, bold_prefix="📈 Valor agregado para colegios sin plataforma", color=COLOR_NAVY_MED, size=8.8)
    p_an2 = c_an2.add_paragraph()
    p_an2.paragraph_format.space_before = Pt(2)
    p_an2.paragraph_format.space_after = Pt(4)
    p_an2.paragraph_format.line_spacing = 1.15
    r_an2 = p_an2.add_run("Si su colegio hoy no cuenta con una solución digital, Skooling aporta beneficios inmediatos y medibles para toda la comunidad:")
    r_an2.font.name = 'Arial'; r_an2.font.size = Pt(7.8)
    
    add_bullet_item(c_an2, "Valor agregado reconocido por familias:", "La comunicación directa y oportuna entre institución y hogares se reconoce como un gran sello de calidad.", size=7.6)
    add_bullet_item(c_an2, "Ahorro directo en papelería e impresiones:", "Boletines digitales en PDF, circulares sin papel y observador en línea.", size=7.6)
    add_bullet_item(c_an2, "Descongestión en secretaría y portería:", "Al consultar en tiempo real notas, agenda y rutas, disminuyen drásticamente llamadas y reclamos.", size=7.6)

    # Sección 8: Garantías Técnicas
    create_section_header(doc, "8. Garantías Técnicas Incluidas en Todos los Planes Anuales", "INFRAESTRUCTURA CLOUD")
    
    tech_tbl = doc.add_table(rows=1, cols=1)
    tech_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tech_tbl.autofit = False
    c_tech = tech_tbl.cell(0, 0)
    c_tech.width = Inches(7.3)
    set_cell_background(c_tech, HEX_NAVY_DARK)
    set_cell_margins(c_tech, top=90, bottom=90, left=120, right=120)
    set_cell_borders(c_tech, top={'color': 'FFD700', 'sz': '4'}, bottom={'color': 'FFD700', 'sz': '4'},
                            left={'color': 'FFD700', 'sz': '16'}, right={'color': 'FFD700', 'sz': '4'})
    
    format_cell_text(c_tech, bold_prefix="GARANTÍAS DE SEGURIDAD Y CONTINUIDAD OPERATIVA", color=RGBColor(255, 215, 0), size=8.5)
    
    # 2 columnas internas dentro de la celda de garantías
    add_bullet_item(c_tech, "⚡ Infraestructura Google Cloud:", "Servidores de alta velocidad y disponibilidad continua sin requerir equipos locales en el colegio.", size=7.6, check_color=RGBColor(255, 215, 0))
    add_bullet_item(c_tech, "🔒 Seguridad y Privacidad de Datos:", "Cifrado de extremo a extremo, control de roles y copias de seguridad automáticas diarias.", size=7.6, check_color=RGBColor(255, 215, 0))
    add_bullet_item(c_tech, "📱 Publicación Oficial en Tiendas:", "Acceso gratuito para acudientes y docentes en App Store (iOS) y Google Play (Android).", size=7.6, check_color=RGBColor(255, 215, 0))
    add_bullet_item(c_tech, "🛡️ Tarifa Congelada 12 Meses:", "Garantía contractual de precio fijo durante la totalidad del año lectivo contratado.", size=7.6, check_color=RGBColor(255, 215, 0))

    # ==========================================================================
    # PÁGINA 4: CONDICIONES, BONIFICACIONES, CRONOGRAMA Y CIERRE
    # ==========================================================================
    doc.add_page_break()

    # Sección 9: Condiciones del Servicio y Sitio Web Opcional
    create_section_header(doc, "9. Condiciones del Servicio Anual y Alcance Complementario", "ACUERDO INSTITUCIONAL")
    
    grid9_tbl = doc.add_table(rows=1, cols=2)
    grid9_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    grid9_tbl.autofit = False
    c_s9_1 = grid9_tbl.cell(0, 0)
    c_s9_2 = grid9_tbl.cell(0, 1)
    c_s9_1.width = Inches(3.55)
    c_s9_2.width = Inches(3.55)
    
    # Contrato Anual
    set_cell_background(c_s9_1, HEX_GOLD_BG)
    set_cell_margins(c_s9_1, top=80, bottom=80, left=100, right=100)
    set_cell_borders(c_s9_1, top={'color': 'C6A15B', 'sz': '4'}, bottom={'color': 'C6A15B', 'sz': '4'},
                             left={'color': 'C6A15B', 'sz': '16'}, right={'color': 'C6A15B', 'sz': '4'})
    
    format_cell_text(c_s9_1, bold_prefix="🔒 Contrato Anual Unificado y Garantía de Servicio", color=COLOR_NAVY_DARK, size=8.5)
    p_con = c_s9_1.add_paragraph()
    p_con.paragraph_format.space_before = Pt(2)
    p_con.paragraph_format.space_after = Pt(4)
    p_con.paragraph_format.line_spacing = 1.15
    r_con = p_con.add_run("El servicio se suscribe bajo un único convenio anual por año lectivo con tarifa fija garantizada, asegurando operación continua e ininterrumpida sin cobros por usuario activo.")
    r_con.font.name = 'Arial'; r_con.font.size = Pt(7.6)
    
    p_con_sub = c_s9_1.add_paragraph()
    p_con_sub.paragraph_format.space_before = Pt(2)
    p_con_sub.paragraph_format.space_after = Pt(0)
    r_csub = p_con_sub.add_run("Las condiciones y fechas de facturación se definen conjuntamente con la institución en la mesa de negociación para ajustarse a su planeación presupuestal.")
    r_csub.font.name = 'Arial'; r_csub.font.size = Pt(7.2); r_csub.font.color.rgb = COLOR_GOLD_DARK; r_csub.font.bold = True

    # Sitio Web Opcional
    set_cell_background(c_s9_2, HEX_LIGHT_BG)
    set_cell_margins(c_s9_2, top=80, bottom=80, left=100, right=100)
    set_cell_borders(c_s9_2, top={'color': HEX_BORDER, 'sz': '4'}, bottom={'color': HEX_BORDER, 'sz': '4'},
                             left={'color': '1E3254', 'sz': '16'}, right={'color': HEX_BORDER, 'sz': '4'})
    
    format_cell_text(c_s9_2, bold_prefix="🌐 Sitio Web Institucional y Dominio .edu (Opcional)", color=COLOR_NAVY_MED, size=8.5)
    p_web = c_s9_2.add_paragraph()
    p_web.paragraph_format.space_before = Pt(2)
    p_web.paragraph_format.space_after = Pt(4)
    p_web.paragraph_format.line_spacing = 1.15
    r_web = p_web.add_run("Si el colegio actualmente no cuenta con un sitio web oficial propio, este componente puede contemplarse dentro de la negociación comercial:")
    r_web.font.name = 'Arial'; r_web.font.size = Pt(7.6)
    
    add_bullet_item(c_s9_2, "Presencia web oficial:", "Posibilidad de incluir el diseño e implementación de la página institucional del colegio.", size=7.4)
    add_bullet_item(c_s9_2, "Gestión de dominio .edu:", "Aura puede encargarse del acompañamiento y gestión del dominio educativo oficial para el colegio.", size=7.4)

    # Sección 10: Bonificaciones Especiales ($0 COP)
    create_section_header(doc, "10. Bonificaciones Especiales Incluidas Sin Costo ($0 COP)", "VALOR AGREGADO INCLUIDO")
    
    bon_tbl = doc.add_table(rows=1, cols=3)
    bon_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    bon_tbl.autofit = False
    
    c_b1 = bon_tbl.cell(0, 0)
    c_b2 = bon_tbl.cell(0, 1)
    c_b3 = bon_tbl.cell(0, 2)
    c_b1.width = Inches(2.4)
    c_b2.width = Inches(2.4)
    c_b3.width = Inches(2.5)
    
    for c_bon, tag_b, title_b, desc_b in [
        (c_b1, "100% BONIFICADO ($0 COP)", "Parametrización y Migración Inicial",
         "Configuración del ambiente del colegio, escalas de evaluación y carga masiva inicial de estudiantes, docentes, cursos, asignaturas y rutas sin costo adicional."),
        (c_b2, "INCLUIDO SIN COSTO", "Capacitación a Todo el Equipo",
         "Jornadas de entrenamiento guiado para directivos, secretaría académica, cuerpo docente, conductores y monitores de transporte escolar."),
        (c_b3, "SOPORTE VIP", "Kit de Lanzamiento para Familias",
         "Piezas digitales e instructivos listos para compartir con acudientes, asegurando una descarga y adopción ágil de la App desde el primer día.")
    ]:
        set_cell_background(c_bon, HEX_LIGHT_BG)
        set_cell_margins(c_bon, top=70, bottom=70, left=80, right=80)
        set_cell_borders(c_bon, top={'color': '2E7D32', 'sz': '8'}, bottom={'color': HEX_BORDER, 'sz': '4'},
                                left={'color': HEX_BORDER, 'sz': '4'}, right={'color': HEX_BORDER, 'sz': '4'})
        
        format_cell_text(c_bon, bold_prefix=tag_b + "\n", color=RGBColor(46, 125, 50), size=7.2)
        p_bt = c_bon.paragraphs[0]
        r_bt = p_bt.add_run(title_b + "\n")
        r_bt.font.name = 'Arial'; r_bt.font.size = Pt(8.2); r_bt.font.bold = True; r_bt.font.color.rgb = COLOR_NAVY_DARK
        r_bd = p_bt.add_run(desc_b)
        r_bd.font.name = 'Arial'; r_bd.font.size = Pt(7.3); r_bd.font.color.rgb = COLOR_TEXT_MAIN

    # Sección 11: Cronograma de Puesta en Marcha
    create_section_header(doc, "11. Cronograma de Puesta en Marcha", "IMPLEMENTACIÓN ÁGIL")
    
    crono_tbl = doc.add_table(rows=4, cols=3)
    crono_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    crono_tbl.autofit = False
    
    cr_widths = [Inches(1.8), Inches(2.2), Inches(3.3)]
    for row in crono_tbl.rows:
        for idx, width in enumerate(cr_widths):
            row.cells[idx].width = width
            set_cell_margins(row.cells[idx], top=50, bottom=50, left=70, right=70)
            
    cr_headers = ["Fase", "Actividad Principal", "Entregable para la Institución"]
    for idx, text in enumerate(cr_headers):
        cell = crono_tbl.cell(0, idx)
        set_cell_background(cell, HEX_NAVY_DARK)
        set_cell_borders(cell, top={'color': HEX_BORDER, 'sz': '4'}, bottom={'color': HEX_BORDER, 'sz': '4'})
        format_cell_text(cell, bold_prefix=text, color=RGBColor(255, 255, 255), size=7.8)

    cr_rows = [
        ("Paso 1 • Fase Inicial", "Demostración en vivo y definición de alcance", 
         "Sesión ejecutiva de 30 minutos con directivas y selección del módulo y rango de estudiantes."),
        ("Paso 2 • Configuración", "Firma de convenio y parametrización Cloud", 
         "Activación del entorno institucional, configuración de escalas y carga masiva de bases de datos."),
        ("Paso 3 • Lanzamiento", "Capacitación y salida en vivo", 
         "Inducción práctica al equipo de profesores y transporte; entrega de accesos a familias e inicio de operación.")
    ]

    for r_idx, (fase, act, ent) in enumerate(cr_rows, start=1):
        c0 = crono_tbl.cell(r_idx, 0)
        c1 = crono_tbl.cell(r_idx, 1)
        c2 = crono_tbl.cell(r_idx, 2)
        
        for cell in [c0, c1, c2]:
            set_cell_borders(cell, top=border_row, bottom=border_row, left=border_row, right=border_row)
            set_cell_background(cell, HEX_WHITE)
            
        format_cell_text(c0, bold_prefix=fase, color=COLOR_NAVY_MED, size=7.5)
        format_cell_text(c1, bold_prefix=act, color=COLOR_TEXT_MAIN, size=7.5)
        format_cell_text(c2, text=ent, color=COLOR_TEXT_MAIN, size=7.5)

    # Sección 12: Bloque de Cierre y Contacto
    doc.add_paragraph().paragraph_format.space_before = Pt(4)
    
    close_tbl = doc.add_table(rows=1, cols=2)
    close_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    close_tbl.autofit = False
    
    c_cl_l = close_tbl.cell(0, 0)
    c_cl_r = close_tbl.cell(0, 1)
    c_cl_l.width = Inches(4.7)
    c_cl_r.width = Inches(2.6)
    
    set_cell_background(c_cl_l, HEX_NAVY_DARK)
    set_cell_background(c_cl_r, HEX_NAVY_DARK)
    set_cell_margins(c_cl_l, top=100, bottom=100, left=140, right=140)
    set_cell_margins(c_cl_r, top=100, bottom=100, left=140, right=140)
    
    border_cl = {'val': 'single', 'sz': '6', 'color': 'C6A15B'}
    set_cell_borders(c_cl_l, top=border_cl, bottom=border_cl, left=border_cl)
    set_cell_borders(c_cl_r, top=border_cl, bottom=border_cl, right=border_cl)
    
    p_cl = c_cl_l.paragraphs[0]
    p_cl.paragraph_format.space_before = Pt(0)
    p_cl.paragraph_format.space_after = Pt(2)
    r_tag_c = p_cl.add_run("RESPALDO EMPRESARIAL\n")
    r_tag_c.font.name = 'Arial'; r_tag_c.font.size = Pt(7.5); r_tag_c.font.bold = True; r_tag_c.font.color.rgb = COLOR_GOLD
    
    r_emp = p_cl.add_run("Aura Soluciones e Innovación S.A.S.\n")
    r_emp.font.name = 'Arial'; r_emp.font.size = Pt(11.0); r_emp.font.bold = True; r_emp.font.color.rgb = RGBColor(255, 255, 255)
    
    r_cldesc = p_cl.add_run("Ingeniería de software colombiana enfocada en soluciones Cloud de alto impacto. Acompañamos a su institución durante todo el año escolar con soporte cercano y evolución continua.")
    r_cldesc.font.name = 'Arial'; r_cldesc.font.size = Pt(7.6); r_cldesc.font.color.rgb = RGBColor(203, 213, 224)
    
    p_cr = c_cl_r.paragraphs[0]
    p_cr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_cr.paragraph_format.space_before = Pt(0)
    p_cr.paragraph_format.space_after = Pt(2)
    
    r_web = p_cr.add_run("🌐 www.aurasi.co\n")
    r_web.font.name = 'Arial'; r_web.font.size = Pt(9.0); r_web.font.bold = True; r_web.font.color.rgb = COLOR_GOLD
    
    r_mail = p_cr.add_run("✉️ contacto@aurasi.co\n\n")
    r_mail.font.name = 'Arial'; r_mail.font.size = Pt(8.2); r_mail.font.color.rgb = RGBColor(255, 255, 255)
    
    r_vig = p_cr.add_run("Vigencia de la oferta: 30 días calendario")
    r_vig.font.name = 'Arial'; r_vig.font.size = Pt(7.2); r_vig.font.italic = True; r_vig.font.color.rgb = RGBColor(148, 163, 184)

    # Guardar documento
    doc.save(output_path)
    print(f"Documento DOCX generado exitosamente en: {output_path}")

if __name__ == "__main__":
    out = "/Users/camilopedraza/Projects/AuraWeb/propuesta/Propuesta_Comercial_Skooling_Aura.docx"
    assets = "/Users/camilopedraza/Projects/AuraWeb/propuesta/presentation_assets"
    generate_commercial_docx(out, assets)
