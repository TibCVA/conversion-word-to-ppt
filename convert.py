#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script de conversion Word -> PPTX, adapté pour un template sans "placeholders"
mais avec un masque contenant des champs texte tels que "PROJECT TITLE", etc.

Hypothèse : dans le template_CVA.pptx, il n'y a pas de slides pré-existantes,
juste un Master + Slide Layouts contenant des formes (shapes) dont le texte
défaut est :
   - "PROJECT TITLE"
   - "CVA Presentation title"
   - "Subtitle"
pour la page de couverture (layout "Diapositive de titre"),

et pour les slides standard (layout "Slide_standard layout"), on a des textes :
   - "Click to edit Master title style"
   - "[Optional subtitle]"
   - "Modifiez les styles du texte du masque"

IMPORTANT :
- On crée les slides via add_slide(<layout>) puis on parcourt slide.shapes
  (au lieu de slide.placeholders).
"""

import sys
import os
from docx import Document
from pptx import Presentation
from pptx.util import Pt

# -------------------------------------------------------------------------
def parse_docx_to_slides(doc_path):
    """
    Lit le document Word et retourne une liste de slides_data.
    slide_data : { "title": str, "subtitle": str, "blocks": [("paragraph", paragraph), ...] }
    """
    doc = Document(doc_path)
    slides = []
    current_slide = None

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # Début d'une nouvelle diapo
        if text.upper().startswith("SLIDE"):
            if current_slide is not None:
                slides.append(current_slide)
            current_slide = {"title": "", "subtitle": "", "blocks": []}
            continue

        # Titre
        if text.startswith("Titre :"):
            if current_slide is not None:
                current_slide["title"] = text[len("Titre :"):].strip()
            continue

        # Sous-titre
        if text.startswith("Sous-titre / Message clé :"):
            if current_slide is not None:
                current_slide["subtitle"] = text[len("Sous-titre / Message clé :"):].strip()
            continue

        # Sinon, c'est un paragraphe de contenu
        if current_slide is not None:
            current_slide["blocks"].append(("paragraph", para))

    # On ajoute la dernière slide si elle existe
    if current_slide is not None:
        slides.append(current_slide)

    return slides

# -------------------------------------------------------------------------
def get_list_type(paragraph):
    """
    Détecte bullet / decimal via le XML.
    """
    xml = paragraph._p.xml
    if 'w:numFmt val="bullet"' in xml:
        return "bullet"
    elif 'w:numFmt val="decimal"' in xml:
        return "decimal"
    return None

def get_indentation_level(paragraph):
    """
    Convertit l'indentation left_indent en un niveau hiérarchique
    """
    indent = paragraph.paragraph_format.left_indent
    if indent is None:
        return 0
    try:
        return int(indent.inches / 0.25)
    except:
        return 0

def add_paragraph_with_runs(text_frame, paragraph, counters):
    """
    Ajoute un paragraphe dans le text_frame. Gère bullet/decimal + runs.
    """
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
            if run.font.size:
                r.font.size = run.font.size
            else:
                r.font.size = Pt(14)
        return p

# -------------------------------------------------------------------------
def clear_shapes_text(slide):
    """
    Efface le texte de tous les shapes possédant un text_frame.
    """
    for shape in slide.shapes:
        if shape.has_text_frame:
            shape.text = ""

def fill_shapes_text(slide, slide_data, slide_index):
    """
    Au lieu d'utiliser "placeholders", on parcourt slide.shapes.
    On détecte :
      - pour la slide 0 : "project title", "cva presentation title", "subtitle"
      - pour les slides standard : "click to edit master title style",
                                  "[optional subtitle]",
                                  "modifiez les styles du texte du masque"
    et on insère title/subtitle/blocks.
    """
    if slide_index == 0:
        # Diapo de couverture
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            txt_ref = shape.text.strip().lower()

            if "project title" in txt_ref:
                shape.text = slide_data["title"]
            elif "cva presentation title" in txt_ref:
                shape.text = slide_data["title"]
            elif "subtitle" in txt_ref:
                shape.text = slide_data["subtitle"]
    else:
        # Diapo standard
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            txt_ref = shape.text.strip().lower()

            # Titre
            if "click to edit master title style" in txt_ref:
                shape.text = slide_data["title"]

            # Sous-titre
            elif "[optional subtitle]" in txt_ref:
                shape.text = slide_data["subtitle"]

            # Contenu
            elif "modifiez les styles du texte du masque" in txt_ref:
                tf = shape.text_frame
                tf.clear()
                counters = {}
                for (block_type, block_para) in slide_data["blocks"]:
                    if block_type == "paragraph":
                        add_paragraph_with_runs(tf, block_para, counters)
                # Pour gérer d'autres types (table, image), rajouter un if block_type == "table": ...
    # Fin fill_shapes_text

def create_ppt_from_docx(input_docx, template_path, output_pptx):
    # on parse le doc
    slides_data = parse_docx_to_slides(input_docx)

    # on ouvre le template
    prs = Presentation(template_path)

    # On identifie les layouts par name
    cover_layout = None
    standard_layout = None
    for layout in prs.slide_layouts:
        if layout.name == "Diapositive de titre":
            cover_layout = layout
        elif layout.name == "Slide_standard layout":
            standard_layout = layout

    # fallback si non trouvé
    if not cover_layout:
        cover_layout = prs.slide_layouts[0]
    if not standard_layout:
        # supposons qu'il s'agit du layout #1
        standard_layout = prs.slide_layouts[1]

    # on génère les slides
    for idx, slide_data in enumerate(slides_data):
        layout = cover_layout if idx == 0 else standard_layout
        new_slide = prs.slides.add_slide(layout)
        clear_shapes_text(new_slide)
        fill_shapes_text(new_slide, slide_data, idx)

    prs.save(output_pptx)
    print("Conversion OK :", output_pptx)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage : python convert.py input.docx output.pptx")
        sys.exit(1)

    input_docx = sys.argv[1]
    output_pptx = sys.argv[2]
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(script_dir, "template_CVA.pptx")

    create_ppt_from_docx(input_docx, template_path, output_pptx)
