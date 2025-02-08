#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script de conversion d'un document Word (.docx) en présentation PowerPoint (.pptx)
en utilisant un template existant (template_CVA.pptx).

Points clés :
- Le document Word contient des balises SLIDE X, Titre : ..., Sous-titre / Message clé : ...
- Le template PPT doit avoir deux layouts :
   1) "Diapositive de titre" pour la slide 0,
   2) "Slide_standard layout" pour les slides 1+,
  avec des zones de texte portant (dans leur texte) :
   - Diapo titre : "PROJECT TITLE", "CVA Presentation title", "Subtitle"
   - Diapo standard : "Click to edit Master title style",
                     "[Optional subtitle]",
                     "Modifiez les styles du texte du masque".
- S'appuie sur python-docx (pour lire Word) et python-pptx (pour manipuler PowerPoint).
- Gère les shapes groupées via un parcours récursif (itération de group shapes).
- Gère les listes à puces, listes numérotées, gras/italique, etc. pour le contenu.

Usage :
  python convert.py input.docx output.pptx

Le template est supposé s'appeler template_CVA.pptx et être dans le même dossier.
"""

import sys
import os
from docx import Document
from pptx import Presentation
from pptx.util import Pt

# ------------------------------------------------------------------------------------
# 1) Parcours récursif des shapes (gestion group-shapes)
# ------------------------------------------------------------------------------------
def iterate_all_shapes(shape_collection):
    """
    Génère chaque shape (et sous-shape si c'est un group) depuis shape_collection.
    """
    for shp in shape_collection:
        yield shp
        # shape_type == 6 => MSO_SHAPE_TYPE.GROUP
        if shp.shape_type == 6:
            for subshp in iterate_all_shapes(shp.shapes):
                yield subshp

# ------------------------------------------------------------------------------------
# 2) Lecture du Word => liste de slides
# ------------------------------------------------------------------------------------
def parse_docx_to_slides(doc_path):
    """
    Retourne une liste de dictionnaires {title, subtitle, blocks},
    où blocks est une liste de (block_type, data).
    """
    doc = Document(doc_path)
    slides_data = []
    current_slide = None

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # On repère "SLIDE X"
        if text.upper().startswith("SLIDE"):
            if current_slide is not None:
                slides_data.append(current_slide)
            current_slide = {"title": "", "subtitle": "", "blocks": []}
            continue

        # On repère "Titre :"
        if text.startswith("Titre :"):
            if current_slide is not None:
                current_slide["title"] = text[len("Titre :"):].strip()
            continue

        # On repère "Sous-titre / Message clé :"
        if text.startswith("Sous-titre / Message clé :"):
            if current_slide is not None:
                current_slide["subtitle"] = text[len("Sous-titre / Message clé :"):].strip()
            continue

        # Sinon, c'est un paragraphe de contenu
        if current_slide is not None:
            current_slide["blocks"].append(("paragraph", para))

    # Ne pas oublier la dernière diapo si elle existe
    if current_slide is not None:
        slides_data.append(current_slide)

    return slides_data

# ------------------------------------------------------------------------------------
# 3) Détection bullet / decimal par le XML docx
# ------------------------------------------------------------------------------------
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
        return int(indent.inches / 0.25)  # 0.25 inch -> 1 niveau
    except:
        return 0

# ------------------------------------------------------------------------------------
# 4) Insertion d'un paragraphe dans la zone de texte PPT
# ------------------------------------------------------------------------------------
def add_paragraph_with_runs(text_frame, paragraph, counters):
    p = text_frame.add_paragraph()
    style_list = get_list_type(paragraph)
    level = get_indentation_level(paragraph)

    if style_list == "bullet":
        # Liste à puces
        p.text = "• " + paragraph.text
        p.level = level
    elif style_list == "decimal":
        # Liste numérotée
        count = counters.get(level, 0) + 1
        counters[level] = count
        p.text = f"{count}. " + paragraph.text
        p.level = level
    else:
        # Paragraphe normal
        for run in paragraph.runs:
            r = p.add_run()
            r.text = run.text
            if run.bold is not None:
                r.font.bold = run.bold
            if run.italic is not None:
                r.font.italic = run.italic
            if run.underline is not None:
                r.font.underline = run.underline
            # Taille
            if run.font.size:
                r.font.size = run.font.size
            else:
                r.font.size = Pt(14)
    return p

# ------------------------------------------------------------------------------------
# 5) Vider le texte des shapes
# ------------------------------------------------------------------------------------
def clear_shapes_text(slide):
    for shp in iterate_all_shapes(slide.shapes):
        if shp.has_text_frame:
            shp.text = ""

# ------------------------------------------------------------------------------------
# 6) Remplir les shapes par titre / sous-titre / contenu
# ------------------------------------------------------------------------------------
def fill_shapes_text(slide, slide_data, slide_index):
    """
    slide_index = 0 => couverture
    slides 1+ => standard
    """
    if slide_index == 0:
        # Couverture
        for shp in iterate_all_shapes(slide.shapes):
            if not shp.has_text_frame:
                continue
            txt_ref = shp.text.strip().lower()
            if "project title" in txt_ref:
                shp.text = slide_data["title"]
            elif "cva presentation title" in txt_ref:
                shp.text = slide_data["title"]
            elif "subtitle" in txt_ref:
                shp.text = slide_data["subtitle"]
    else:
        # Slides standard
        for shp in iterate_all_shapes(slide.shapes):
            if not shp.has_text_frame:
                continue
            txt_ref = shp.text.strip().lower()
            if "click to edit master title style" in txt_ref:
                shp.text = slide_data["title"]
            elif "[optional subtitle]" in txt_ref:
                shp.text = slide_data["subtitle"]
            elif "modifiez les styles du texte du masque" in txt_ref:
                tf = shp.text_frame
                tf.clear()
                counters = {}
                for (block_type, block_para) in slide_data["blocks"]:
                    if block_type == "paragraph":
                        add_paragraph_with_runs(tf, block_para, counters)

# ------------------------------------------------------------------------------------
# 7) Fonction principale
# ------------------------------------------------------------------------------------
def create_ppt_from_docx(input_docx, template_pptx, output_pptx):
    # 1. On parse le Word
    slides_data = parse_docx_to_slides(input_docx)
    # 2. On ouvre le template
    prs = Presentation(template_pptx)

    # 3. Chercher les layouts
    cover_layout = None
    standard_layout = None
    for layout in prs.slide_layouts:
        if layout.name == "Diapositive de titre":
            cover_layout = layout
        elif layout.name == "Slide_standard layout":
            standard_layout = layout

    # Fallback
    if not cover_layout:
        cover_layout = prs.slide_layouts[0]
    if not standard_layout:
        standard_layout = prs.slide_layouts[1]

    # 4. Générer les slides
    for idx, slide_data in enumerate(slides_data):
        layout = cover_layout if idx == 0 else standard_layout
        slide = prs.slides.add_slide(layout)

        # Vider le texte
        clear_shapes_text(slide)
        # Remplir
        fill_shapes_text(slide, slide_data, idx)

    # 5. Sauver
    prs.save(output_pptx)
    print("Conversion terminée :", output_pptx)

# ------------------------------------------------------------------------------------
# 8) Point d'entrée
# ------------------------------------------------------------------------------------
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage : python convert.py input.docx output.pptx")
        sys.exit(1)

    input_docx = sys.argv[1]
    output_pptx = sys.argv[2]
    # Le template doit être dans le même dossier : template_CVA.pptx
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_file = os.path.join(script_dir, "template_CVA.pptx")

    create_ppt_from_docx(input_docx, template_file, output_pptx)
