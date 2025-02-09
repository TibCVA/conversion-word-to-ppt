#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Conversion d'un document Word (.docx) en présentation PowerPoint (.pptx)
en utilisant un template existant (template_CVA.pptx) comme arrière-plan.
Le document Word doit être structuré de la manière suivante :
  • Chaque slide commence par une ligne "SLIDE X"
  • Une ligne "Titre :" indique le titre de la slide
  • Une ligne "Sous-titre / Message clé :" indique le sous-titre
  • Le reste (paragraphes, listes, tableaux) constitue le contenu de la slide

Le script crée pour chaque slide une diapositive à partir d'un layout "Blank" du template,
et y ajoute trois zones de texte positionnées précisément selon ces coordonnées (en pixels, converties en pouces) :

  title_zone:    { x: 76, y: 35, width: 1382, height: 70 }
  subtitle_zone: { x: 76, y: 119, width: 1382, height: 56 }
  content_zone:  { x: 76, y: 189, width: 1382, height: 425 }

Les styles forcés sont :
  - Titre : Arial, taille 22 pts en gras (auto-ajusté entre 22 et 16 pts)
  - Sous-titre : Arial, taille 18 pts non gras (auto-ajusté entre 18 et 14 pts)
  - Contenu : Arial, taille 11 pts (auto-ajusté entre 11 et 9 pts)
  
Les paragraphes conservent les puces, numérotations et retraits.
Les tableaux sont insérés en reproduisant leur contenu cellule par cellule en Arial 10 pts (la première ligne en gras).

Usage :
  python convert_new.py input.docx output.pptx

Le fichier template_CVA.pptx doit se trouver dans le même dossier que ce script.
"""

import sys
import os
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt

# Conversion de pixels en pouces (1 pouce = 96 pixels)
def px_to_inch(px):
    return px / 96.0

# Coordonnées des zones converties en pouces
TITLE_ZONE = {
    "x": px_to_inch(76),
    "y": px_to_inch(35),
    "width": px_to_inch(1382),
    "height": px_to_inch(70)
}
SUBTITLE_ZONE = {
    "x": px_to_inch(76),
    "y": px_to_inch(119),
    "width": px_to_inch(1382),
    "height": px_to_inch(56)
}
CONTENT_ZONE = {
    "x": px_to_inch(76),
    "y": px_to_inch(189),
    "width": px_to_inch(1382),
    "height": px_to_inch(425)
}

# ------------------------------------------------------------------------------
# Itération sur les éléments de niveau bloc (paragraphes et tableaux) dans le Word
# ------------------------------------------------------------------------------
def iter_block_items(parent):
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    parent_elm = parent.element
    for child in parent_elm.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, parent)
        elif child.tag == qn('w:tbl'):
            yield Table(child, parent)

# ------------------------------------------------------------------------------
# Extraction du contenu du Word en slides
# ------------------------------------------------------------------------------
def parse_docx_to_slides(doc_path):
    doc = Document(doc_path)
    slides_data = []
    current_slide = None

    for block in iter_block_items(doc):
        if block.__class__.__name__ == "Paragraph":
            text = block.text.strip()
            if not text:
                continue
            if text.upper().startswith("SLIDE"):
                if current_slide is not None:
                    slides_data.append(current_slide)
                current_slide = {"title": "", "subtitle": "", "blocks": []}
                continue
            if text.startswith("Titre :"):
                if current_slide is not None:
                    current_slide["title"] = text[len("Titre :"):].strip()
                continue
            if text.startswith("Sous-titre / Message clé :"):
                if current_slide is not None:
                    current_slide["subtitle"] = text[len("Sous-titre / Message clé :"):].strip()
                continue
            if current_slide is not None:
                current_slide["blocks"].append(("paragraph", block))
        else:
            # Traitement des tableaux
            if current_slide is not None:
                current_slide["blocks"].append(("table", block))
    if current_slide is not None:
        slides_data.append(current_slide)
    print("Contenu extrait du Word:")
    for i, slide in enumerate(slides_data):
        print(f"Slide {i}: Titre='{slide['title']}', Sous-titre='{slide['subtitle']}', Nb blocs={len(slide['blocks'])}")
    return slides_data

# ------------------------------------------------------------------------------
# Gestion du formatage pour les paragraphes
# ------------------------------------------------------------------------------
def get_list_type(paragraph):
    xml = paragraph._p.xml
    if 'w:numFmt val="bullet"' in xml:
        return "bullet"
    elif 'w:numFmt val="decimal"' in xml:
        return "decimal"
    return None

def get_indentation_level(paragraph):
    indent = paragraph.paragraph_format.left_indent
    if indent is None:
        return 0
    try:
        return int(indent.inches / 0.25)
    except:
        return 0

def add_paragraph_with_runs(text_frame, paragraph, counters):
    new_p = text_frame.add_paragraph()
    new_p.level = get_indentation_level(paragraph)
    style = get_list_type(paragraph)
    if style == "bullet":
        r = new_p.add_run()
        r.text = "• "
        r.font.name = "Arial"
        r.font.bold = True
        for run in paragraph.runs:
            r = new_p.add_run()
            r.text = run.text
            r.font.name = "Arial"
            r.font.bold = run.bold if run.bold is not None else False
            r.font.italic = run.italic if run.italic is not None else False
            r.font.underline = run.underline if run.underline is not None else False
            r.font.size = run.font.size if run.font.size else Pt(14)
    elif style == "decimal":
        count = counters.get(new_p.level, 0) + 1
        counters[new_p.level] = count
        r = new_p.add_run()
        r.text = f"{count}. "
        r.font.name = "Arial"
        r.font.bold = True
        for run in paragraph.runs:
            r = new_p.add_run()
            r.text = run.text
            r.font.name = "Arial"
            r.font.bold = run.bold if run.bold is not None else False
            r.font.italic = run.italic if run.italic is not None else False
            r.font.underline = run.underline if run.underline is not None else False
            r.font.size = run.font.size if run.font.size else Pt(14)
    else:
        for run in paragraph.runs:
            r = new_p.add_run()
            r.text = run.text
            r.font.name = "Arial"
            r.font.bold = run.bold if run.bold is not None else False
            r.font.italic = run.italic if run.italic is not None else False
            r.font.underline = run.underline if run.underline is not None else False
            r.font.size = run.font.size if run.font.size else Pt(14)
    return new_p

# ------------------------------------------------------------------------------
# Insertion d'un tableau dans la slide PowerPoint
# ------------------------------------------------------------------------------
def insert_table_in_slide(slide, table, left, top, width, height):
    rows = len(table.rows)
    cols = len(table.columns)
    shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    ppt_table = shape.table
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            cell_text = "\n".join(p.text for p in cell.paragraphs)
            ppt_table.cell(i, j).text = cell_text
            for paragraph in ppt_table.cell(i, j).text_frame.paragraphs:
                for r in paragraph.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(10)
                    if i == 0:
                        r.font.bold = True
    return shape

# ------------------------------------------------------------------------------
# Insertion des blocs (paragraphes et tableaux) dans la zone de contenu
# ------------------------------------------------------------------------------
def add_content_blocks(slide, blocks, zone):
    """
    Insère les blocs (paragraphes et tableaux) dans la zone de contenu définie par 'zone'.
    Commence à zone['y'] et insère chaque bloc verticalement avec un espacement fixe.
    """
    current_top = zone["y"]
    spacing = Inches(0.2)
    for block_type, block in blocks:
        if block_type == "paragraph":
            height = Inches(0.5)  # hauteur approximative pour un paragraphe
            tb = slide.shapes.add_textbox(zone["x"], current_top, zone["width"], height)
            tf = tb.text_frame
            tf.clear()
            counters = {}
            add_paragraph_with_runs(tf, block, counters)
            try:
                tf.fit_text(max_size=11, min_size=9)
            except Exception as e:
                print("Erreur fit_text pour paragraphe:", e)
            current_top += height + spacing
        elif block_type == "table":
            rows = len(block.rows)
            height = Inches(0.3) * rows  # estimation
            insert_table_in_slide(slide, block, left=zone["x"], top=current_top, width=zone["width"], height=height)
            current_top += height + spacing

# ------------------------------------------------------------------------------
# Création d'une slide (pour toutes les slides, un seul template)
# ------------------------------------------------------------------------------
def add_slide_with_text(prs, slide_data):
    """
    Crée une diapositive (à partir d'un layout Blank du template) et y ajoute trois zones de texte
    positionnées selon les coordonnées définies par TITLE_ZONE, SUBTITLE_ZONE et CONTENT_ZONE.
    """
    blank_layout = None
    for layout in prs.slide_layouts:
        if "Blank" in layout.name:
            blank_layout = layout
            break
    if blank_layout is None:
        blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    
    # Zone de titre
    title_box = slide.shapes.add_textbox(TITLE_ZONE["x"], TITLE_ZONE["y"], TITLE_ZONE["width"], TITLE_ZONE["height"])
    title_tf = title_box.text_frame
    title_tf.text = slide_data["title"]
    for p in title_tf.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"
            r.font.bold = True
    title_tf.fit_text(max_size=22, min_size=16)
    
    # Zone de sous-titre
    subtitle_box = slide.shapes.add_textbox(SUBTITLE_ZONE["x"], SUBTITLE_ZONE["y"], SUBTITLE_ZONE["width"], SUBTITLE_ZONE["height"])
    subtitle_tf = subtitle_box.text_frame
    subtitle_tf.text = slide_data["subtitle"]
    for p in subtitle_tf.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"
            r.font.bold = False
    subtitle_tf.fit_text(max_size=18, min_size=14)
    
    # Zone de contenu : insertion des blocs (paragraphes et tableaux)
    add_content_blocks(slide, slide_data["blocks"], CONTENT_ZONE)
    
    return slide

# ------------------------------------------------------------------------------
# Fonction principale de conversion
# ------------------------------------------------------------------------------
def create_ppt_from_docx(input_docx, template_pptx, output_pptx):
    slides_data = parse_docx_to_slides(input_docx)
    prs = Presentation(template_pptx)
    if slides_data:
        for slide_data in slides_data:
            add_slide_with_text(prs, slide_data)
    else:
        print("Aucune slide trouvée dans le document Word.")
    prs.save(output_pptx)
    print("Conversion terminée :", output_pptx)

# ------------------------------------------------------------------------------
# Point d'entrée
# ------------------------------------------------------------------------------
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage : python convert_new.py input.docx output.pptx")
        sys.exit(1)
    input_docx = sys.argv[1]
    output_pptx = sys.argv[2]
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_file = os.path.join(script_dir, "template_CVA.pptx")
    create_ppt_from_docx(input_docx, template_file, output_pptx)