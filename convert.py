#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Convertit un fichier Word en une présentation PowerPoint à l'aide d'un template.
Usage : python convert.py input.docx output.pptx

Le fichier Word doit être structuré ainsi :
  - Chaque diapositive commence par une ligne "SLIDE X" (ex. "SLIDE 1").
  - Une ligne "Titre :" définit le titre.
  - Une ligne "Sous-titre / Message clé :" définit le sous-titre.
  - Le reste du contenu (paragraphes, tableaux) constitue le corps.
  
La première diapositive (SLIDE 0) utilisera le layout de couverture (layout index 0) du template.
Les diapositives suivantes utiliseront le layout standard (layout index 1).

Ce script n'utilise pas de placeholders internes du template : il ajoute des zones de texte avec des positions fixes.
"""

import sys
import os
import docx
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt

def iter_block_items(parent):
    """
    Itère sur les éléments blocs (paragraphes et tableaux) d'un document dans l'ordre.
    """
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.table import _Cell, Table
    from docx.text.paragraph import Paragraph

    if isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        parent_elm = parent.element

    for child in parent_elm.iterchildren():
        if child.tag.endswith('}p'):
            yield Paragraph(child, parent)
        elif child.tag.endswith('}tbl'):
            yield Table(child, parent)

def parse_docx_to_slides(doc_path):
    """
    Parcourt le document Word et découpe le contenu en slides.
    Chaque slide est un dictionnaire avec :
      - "title" : texte après "Titre :"
      - "subtitle" : texte après "Sous-titre / Message clé :"
      - "blocks" : liste d'éléments (tuple ("paragraph", objet) ou ("table", objet))
    """
    doc = Document(doc_path)
    slides = []
    current_slide = None
    for block in iter_block_items(doc):
        if isinstance(block, docx.text.paragraph.Paragraph):
            txt = block.text.strip()
            if not txt:
                continue
            # Début d'une nouvelle slide
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
        elif isinstance(block, docx.table.Table):
            if current_slide is not None:
                current_slide["blocks"].append(("table", block))
    if current_slide is not None:
        slides.append(current_slide)
    return slides

def add_paragraph(text_frame, paragraph):
    """
    Ajoute un paragraphe dans le text_frame en copiant les runs pour conserver le formatage.
    """
    p = text_frame.add_paragraph()
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

def add_table(slide, table, left, top, width, height):
    """
    Ajoute un tableau sur la diapositive en copiant le texte de chaque cellule.
    """
    rows = len(table.rows)
    cols = len(table.columns)
    shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    ppt_table = shape.table
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            ppt_table.cell(i, j).text = cell.text
    return shape

def create_ppt_from_docx(doc_path, template_path, output_path):
    """
    Crée une présentation PowerPoint à partir d'un document Word et d'un template.
    Pour la première diapositive, utilise le layout index 0 ; pour les suivantes, le layout index 1.
    Ajoute manuellement des zones de texte pour le titre, le sous-titre et le contenu.
    """
    slides_data = parse_docx_to_slides(doc_path)
    prs = Presentation(template_path)
    
    # Sélection des layouts : index 0 pour la première slide, index 1 pour les autres.
    cover_layout = prs.slide_layouts[0]
    standard_layout = prs.slide_layouts[1]
    
    for idx, slide_data in enumerate(slides_data):
        layout = cover_layout if idx == 0 else standard_layout
        slide = prs.slides.add_slide(layout)
        
        # Ajout d'une zone de texte pour le titre
        title_left, title_top, title_width, title_height = Inches(0.5), Inches(0.5), Inches(9), Inches(1)
        title_box = slide.shapes.add_textbox(title_left, title_top, title_width, title_height)
        title_box.text_frame.clear()
        title_box.text_frame.text = slide_data["title"]
        
        # Ajout d'une zone de texte pour le sous-titre
        subtitle_left, subtitle_top, subtitle_width, subtitle_height = Inches(0.5), Inches(1.5), Inches(9), Inches(1)
        subtitle_box = slide.shapes.add_textbox(subtitle_left, subtitle_top, subtitle_width, subtitle_height)
        subtitle_box.text_frame.clear()
        subtitle_box.text_frame.text = slide_data["subtitle"]
        
        # Ajout d'une zone de texte pour le contenu
        content_left, content_top, content_width, content_height = Inches(0.5), Inches(2.5), Inches(9), Inches(4)
        content_box = slide.shapes.add_textbox(content_left, content_top, content_width, content_height)
        content_frame = content_box.text_frame
        content_frame.clear()
        for block_type, block_obj in slide_data["blocks"]:
            if block_type == "paragraph":
                add_paragraph(content_frame, block_obj)
            elif block_type == "table":
                # Ajoute le tableau en dehors de la zone de contenu, à une position fixe (modifiable si nécessaire)
                table_left, table_top = Inches(0.5), Inches(7)
                table_width, table_height = Inches(9), Inches(2)
                add_table(slide, block_obj, table_left, table_top, table_width, table_height)
                content_frame.add_paragraph()
    
    prs.save(output_path)
    print("Conversion terminée :", output_path)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python convert.py input.docx output.pptx")
        sys.exit(1)
    input_docx = sys.argv[1]
    output_pptx = sys.argv[2]
    # Le template se trouve dans le même dossier que ce script, nommé "template_CVA.pptx"
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template_CVA.pptx")
    create_ppt_from_docx(input_docx, template_path, output_pptx)
