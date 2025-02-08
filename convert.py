#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Conversion d'un document Word en présentation PowerPoint en insérant le contenu
dans les placeholders du template.

Usage : python convert.py input.docx output.pptx

Structure attendue du Word (input.docx) :
  • Chaque slide commence par une ligne "SLIDE X" (ex. "SLIDE 1")
  • Une ligne "Titre :" définit le titre
  • Une ligne "Sous-titre / Message clé :" définit le sous-titre
  • Le reste (paragraphes et tableaux) constitue le contenu

Template PPT :
  • Pour SLIDE 0 (couverture) : placeholders avec textes par défaut "PROJECT TITLE", "CVA Presentation title", "Subtitle" (et "Date" qui reste inchangé)
  • Pour SLIDES 1+ (standard) : placeholders avec textes par défaut "Click to edit Master title style" (titre), "[Optional subtitle]" (sous-titre) et un placeholder BODY dont le texte par défaut commence par "Modifiez les styles du texte du masque..."
Le script insère le contenu dans ces zones en préservant la mise en forme (listes, gras, souligné) et en insérant les tableaux en dessous du texte.
"""

import sys
import os
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt

# Fonction pour itérer sur tous les blocs (paragraphes et tableaux) dans l'ordre
def iter_block_items(parent):
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    parent_element = parent.element
    for child in parent_element.iterchildren():
        if child.tag.endswith('}p'):
            yield Paragraph(child, parent)
        elif child.tag.endswith('}tbl'):
            yield Table(child, parent)

# Extraction des blocs (paragraphes et tableaux) en respectant l'ordre
def parse_docx_to_slides(doc_path):
    doc = Document(doc_path)
    slides = []
    current_slide = None
    for block in iter_block_items(doc):
        from docx.text.paragraph import Paragraph
        from docx.table import Table
        if isinstance(block, Paragraph):
            txt = block.text.strip()
            if not txt:
                continue
            if txt.upper().startswith("SLIDE"):
                if current_slide is not None:
                    slides.append(current_slide)
                current_slide = {"title": "", "subtitle": "", "blocks": []}
                continue
            if txt.startswith("Titre :"):
                if current_slide is not None:
                    current_slide["title"] = txt[len("Titre :"):].strip()
                continue
            if txt.startswith("Sous-titre / Message clé :"):
                if current_slide is not None:
                    current_slide["subtitle"] = txt[len("Sous-titre / Message clé :"):].strip()
                continue
            if current_slide is not None:
                current_slide["blocks"].append(("paragraph", block))
        elif isinstance(block, Table):
            if current_slide is not None:
                current_slide["blocks"].append(("table", block))
    if current_slide is not None:
        slides.append(current_slide)
    return slides

# Ajout d'un paragraphe dans un text frame en préservant le formatage et la numérotation
def add_paragraph_with_runs(text_frame, paragraph, counters):
    p = text_frame.add_paragraph()
    list_type = get_list_type(paragraph)
    level = get_indentation_level(paragraph)
    if list_type == "bullet":
        p.text = "• " + paragraph.text
        p.level = level
        return p
    elif list_type == "decimal":
        count = counters.get(level, 0) + 1
        counters[level] = count
        p.text = f"{count}. " + paragraph.text
        p.level = level
        return p
    else:
        for run in paragraph.runs:
            r = p.add_run()
            r.text = run.text
            if run.bold is not None:
                r.font.bold = run.bold
            if run.italic is not None:
                r.font.italic = run.italic
            if run.underline is not None:
                r.font.underline = run.underline
            r.font.size = run.font.size if run.font.size else Pt(14)
        return p

# Détection du type de liste à partir du XML du paragraphe
def get_list_type(paragraph):
    xml = paragraph._p.xml
    if 'w:numFmt val="bullet"' in xml:
        return "bullet"
    elif 'w:numFmt val="decimal"' in xml:
        return "decimal"
    else:
        return None

def get_indentation_level(paragraph):
    indent = paragraph.paragraph_format.left_indent
    if indent is None:
        return 0
    try:
        return int(indent.inches / 0.25)
    except Exception:
        return 0

# Insertion d'un tableau dans la diapo, dans une zone définie par (left, top, width, height)
def insert_table(slide, table, left, top, width, height):
    rows = len(table.rows)
    cols = len(table.columns)
    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    ppt_table = table_shape.table
    # Optionnel : copier la largeur des colonnes depuis Word
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            cell_text = "\n".join(p.text for p in cell.paragraphs)
            ppt_table.cell(i, j).text = cell_text
    return table_shape

# Remplissage des placeholders de la diapositive avec le contenu extrait
def fill_placeholders(slide, slide_data, slide_index):
    if slide_index == 0:
        # Pour la couverture, on suppose que le placeholder idx 0 est le titre et idx 1 est le sous-titre.
        for shape in slide.placeholders:
            if not shape.has_text_frame:
                continue
            idx = shape.placeholder_format.idx
            if idx == 0:
                shape.text = slide_data["title"]
            elif idx == 1:
                shape.text = slide_data["subtitle"]
    else:
        # Pour les slides standard, on suppose que idx 0 est le titre et idx 1 est le contenu.
        for shape in slide.placeholders:
            if not shape.has_text_frame:
                continue
            idx = shape.placeholder_format.idx
            if idx == 0:
                shape.text = slide_data["title"]
            elif idx == 1:
                tf = shape.text_frame
                tf.clear()
                # Ajouter le sous-titre en premier si présent
                if slide_data["subtitle"]:
                    p_sub = tf.add_paragraph()
                    p_sub.text = slide_data["subtitle"]
                counters = {}
                # Séparer les blocs en texte et tableaux
                text_blocks = []
                table_blocks = []
                for block_type, block in slide_data["blocks"]:
                    if block_type == "paragraph":
                        text_blocks.append(block)
                    elif block_type == "table":
                        table_blocks.append(block)
                for para in text_blocks:
                    add_paragraph_with_runs(tf, para, counters)
                # Après le texte, insérer chaque tableau en dessous du placeholder BODY
                # On calcule la position en fonction du placeholder BODY
                left = shape.left
                top = shape.top + shape.height + Inches(0.2)
                width = shape.width
                for tbl in table_blocks:
                    # Estimer la hauteur en fonction du nombre de lignes, par ex. 0.8 inch par ligne
                    height = Inches(0.8 * len(tbl.rows))
                    insert_table(slide, tbl, left, top, width, height)
                    top += height + Inches(0.2)

def clear_all_placeholders(slide):
    for shape in slide.placeholders:
        try:
            shape.text = ""
        except Exception:
            pass

def create_ppt_from_docx(doc_path, template_path, output_path):
    slides_data = parse_docx_to_slides(doc_path)
    prs = Presentation(template_path)
    for idx, slide_data in enumerate(slides_data):
        layout = prs.slide_layouts[0] if idx == 0 else prs.slide_layouts[1]
        slide = prs.slides.add_slide(layout)
        clear_all_placeholders(slide)
        fill_placeholders(slide, slide_data, idx)
    prs.save(output_path)
    print("Conversion terminée :", output_path)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python convert.py input.docx output.pptx")
        sys.exit(1)
    input_docx = sys.argv[1]
    output_pptx = sys.argv[2]
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template_CVA.pptx")
    create_ppt_from_docx(input_docx, template_path, output_pptx)
