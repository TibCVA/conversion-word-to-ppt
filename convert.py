#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

def px_to_inch(px):
    return px / 96.0

# Définition des zones en pouces
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

def parse_docx_to_slides(doc_path):
    doc = Document(doc_path)
    slides_data = []
    current_slide = None
    
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
            
        if text.upper().startswith("SLIDE"):
            if current_slide is not None:
                slides_data.append(current_slide)
            current_slide = {"title": "", "subtitle": "", "content": [], "tables": []}
            
        elif text.startswith("Titre :") and current_slide is not None:
            current_slide["title"] = text[len("Titre :"):].strip()
            
        elif text.startswith("Sous-titre / Message clé :") and current_slide is not None:
            current_slide["subtitle"] = text[len("Sous-titre / Message clé :"):].strip()
            
        elif current_slide is not None:
            current_slide["content"].append(paragraph)
    
    if current_slide is not None:
        slides_data.append(current_slide)
    
    # Ajout des tables aux slides
    current_slide_idx = 0
    for table in doc.tables:
        while current_slide_idx < len(slides_data):
            if any(p.text.upper().startswith("SLIDE") for p in table.rows[0].cells[0].paragraphs):
                current_slide_idx += 1
            else:
                slides_data[current_slide_idx]["tables"].append(table)
                break
    
    return slides_data

def create_textbox_with_text(slide, text, left, top, width, height, font_name="Arial", 
                           font_size=11, bold=False, auto_fit_size=None, margin_left=0.1):
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    text_frame.word_wrap = True
    text_frame.margin_left = Inches(margin_left)
    text_frame.margin_right = Inches(0.1)
    text_frame.vertical_anchor = MSO_ANCHOR.TOP
    
    paragraph = text_frame.paragraphs[0]
    paragraph.text = text
    paragraph.font.name = font_name
    paragraph.font.size = Pt(font_size)
    paragraph.font.bold = bold
    paragraph.alignment = PP_ALIGN.LEFT
    
    if auto_fit_size:
        try:
            text_frame.fit_text(max_size=auto_fit_size)
        except:
            pass  # Silently fail if auto-fit doesn't work
            
    return textbox

def add_table_to_slide(slide, table, left, top, width, height):
    rows = len(table.rows)
    cols = len(table.columns)
    
    shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    tbl = shape.table
    
    # Copie du contenu et formatage
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            target_cell = tbl.cell(i, j)
            source_text = " ".join(paragraph.text for paragraph in cell.paragraphs)
            target_cell.text = source_text
            
            # Application du style
            paragraph = target_cell.text_frame.paragraphs[0]
            paragraph.font.name = "Arial"
            paragraph.font.size = Pt(10)
            if i == 0:  # Première ligne en gras
                paragraph.font.bold = True
    
    return shape

def create_slide(prs, slide_data):
    # Recherche du layout Blank
    blank_layout = None
    for layout in prs.slide_layouts:
        if layout.name == "Blank":
            blank_layout = layout
            break
    if not blank_layout:
        blank_layout = prs.slide_layouts[6]  # Layout par défaut si Blank non trouvé
    
    slide = prs.slides.add_slide(blank_layout)
    
    # Ajout du titre
    title_box = create_textbox_with_text(
        slide, slide_data["title"],
        TITLE_ZONE["x"], TITLE_ZONE["y"],
        TITLE_ZONE["width"], TITLE_ZONE["height"],
        font_size=22, bold=True, auto_fit_size=22
    )
    
    # Ajout du sous-titre
    subtitle_box = create_textbox_with_text(
        slide, slide_data["subtitle"],
        SUBTITLE_ZONE["x"], SUBTITLE_ZONE["y"],
        SUBTITLE_ZONE["width"], SUBTITLE_ZONE["height"],
        font_size=18, auto_fit_size=18
    )
    
    # Position initiale pour le contenu
    current_y = CONTENT_ZONE["y"]
    
    # Ajout du contenu texte
    for paragraph in slide_data["content"]:
        if current_y >= (CONTENT_ZONE["y"] + CONTENT_ZONE["height"]):
            break
            
        height = Inches(0.3)  # Hauteur par défaut pour chaque paragraphe
        
        text_box = create_textbox_with_text(
            slide, paragraph.text,
            CONTENT_ZONE["x"], current_y,
            CONTENT_ZONE["width"], height,
            font_size=11, auto_fit_size=11
        )
        
        # Gestion des puces et numérotation
        if hasattr(paragraph, "style") and paragraph.style.name.startswith("List"):
            text_frame = text_box.text_frame
            p = text_frame.paragraphs[0]
            p.level = int(paragraph.style.name[-1]) - 1 if paragraph.style.name[-1].isdigit() else 0
        
        current_y += height + Inches(0.1)  # Espacement entre paragraphes
    
    # Ajout des tableaux
    for table in slide_data["tables"]:
        if current_y >= (CONTENT_ZONE["y"] + CONTENT_ZONE["height"]):
            break
            
        table_height = Inches(len(table.rows) * 0.3)
        add_table_to_slide(
            slide, table,
            CONTENT_ZONE["x"], current_y,
            CONTENT_ZONE["width"], table_height
        )
        current_y += table_height + Inches(0.2)
    
    return slide

def create_ppt_from_docx(input_docx, template_pptx, output_pptx):
    # Extraction des données du Word
    slides_data = parse_docx_to_slides(input_docx)
    if not slides_data:
        print("Aucune slide trouvée dans le document Word.")
        return
    
    # Création de la présentation
    prs = Presentation(template_pptx)
    
    # Suppression des slides existantes
    xml_slides = prs.slides._sldIdLst
    slides_to_remove = list(xml_slides)
    for sld in slides_to_remove:
        xml_slides.remove(sld)
    
    # Création des nouvelles slides
    for slide_data in slides_data:
        create_slide(prs, slide_data)
    
    # Sauvegarde
    prs.save(output_pptx)
    print(f"Conversion terminée: {len(slides_data)} slides créées dans {output_pptx}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python convert.py input.docx output.pptx")
        sys.exit(1)
        
    input_docx = sys.argv[1]
    output_pptx = sys.argv[2]
    template_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template_CVA.pptx")
    
    create_ppt_from_docx(input_docx, template_file, output_pptx)