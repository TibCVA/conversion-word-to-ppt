#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Conversion d'un document Word (.docx) en présentation PowerPoint (.pptx)
en utilisant un template existant (template_CVA.pptx).

Structure attendue du Word :
  • Chaque diapositive commence par une ligne "SLIDE X"
  • Une ligne "Titre :" indique le titre
  • Une ligne "Sous-titre / Message clé :" indique le sous-titre
  • Le reste (paragraphes, listes, …) constitue le contenu

Le template PPT doit comporter deux layouts :
  • "Diapositive de titre" pour la slide de couverture, avec au moins :
       - Un placeholder de type TITLE (pour "PROJECT TITLE" ou "CVA Presentation title")
       - Un placeholder de type SUBTITLE (pour "Subtitle")
  • "Slide_standard layout" pour les slides standards, avec au moins :
       - Un placeholder de type TITLE (pour "Click to edit Master title style")
       - Un placeholder de type SUBTITLE (pour "[Optional subtitle]")
       - Un placeholder de type BODY (pour "Modifiez les styles du texte du masque" et le contenu)

Usage :
  python convert.py input.docx output.pptx
"""

import sys
import os
from docx import Document
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import PP_PLACEHOLDER

# ------------------------------------------------------------------------------------
# Parcours récursif des shapes (gère aussi les group-shapes)
# ------------------------------------------------------------------------------------
def iterate_all_shapes(shape_collection):
    for shp in shape_collection:
        yield shp
        if shp.shape_type == 6:  # 6 = GROUP
            for subshp in iterate_all_shapes(shp.shapes):
                yield subshp

# ------------------------------------------------------------------------------------
# Extraction du contenu Word en liste de slides
# ------------------------------------------------------------------------------------
def parse_docx_to_slides(doc_path):
    doc = Document(doc_path)
    slides_data = []
    current_slide = None

    for para in doc.paragraphs:
        text = para.text.strip()
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
            current_slide["blocks"].append(("paragraph", para))

    if current_slide is not None:
        slides_data.append(current_slide)

    return slides_data

# ------------------------------------------------------------------------------------
# Gestion des listes (bullet / decimal) et insertion de runs
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
        return int(indent.inches / 0.25)  # 0.25 inch par niveau
    except:
        return 0

def add_paragraph_with_runs(text_frame, paragraph, counters):
    new_p = text_frame.add_paragraph()
    style_list = get_list_type(paragraph)
    level = get_indentation_level(paragraph)

    if style_list == "bullet":
        new_p.text = "• " + paragraph.text
        new_p.level = level
    elif style_list == "decimal":
        count = counters.get(level, 0) + 1
        counters[level] = count
        new_p.text = f"{count}. " + paragraph.text
        new_p.level = level
    else:
        for run in paragraph.runs:
            r = new_p.add_run()
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
    return new_p

# ------------------------------------------------------------------------------------
# Vider le texte des shapes d'une slide
# ------------------------------------------------------------------------------------
def clear_shapes_text(slide):
    for shp in iterate_all_shapes(slide.shapes):
        if shp.has_text_frame:
            try:
                shp.text_frame.clear()
            except Exception:
                # Si clear() ne fonctionne pas, on tente de définir le texte vide
                shp.text = ""

# ------------------------------------------------------------------------------------
# Remplir les zones de texte d'une slide en fonction du type de placeholder
# ------------------------------------------------------------------------------------
def fill_shapes_text(slide, slide_data, slide_index):
    if slide_index == 0:
        # Slide de couverture : on met à jour les placeholders TITLE et SUBTITLE
        for shp in iterate_all_shapes(slide.shapes):
            if not shp.has_text_frame:
                continue
            if shp.placeholder_format is not None:
                if shp.placeholder_format.type == PP_PLACEHOLDER.TITLE:
                    shp.text_frame.clear()
                    shp.text = slide_data["title"]
                elif shp.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE:
                    shp.text_frame.clear()
                    shp.text = slide_data["subtitle"]
            else:
                # Fallback par recherche textuelle
                txt_ref = shp.text.strip().lower()
                if "project title" in txt_ref or "cva presentation title" in txt_ref:
                    shp.text_frame.clear()
                    shp.text = slide_data["title"]
                elif "subtitle" in txt_ref:
                    shp.text_frame.clear()
                    shp.text = slide_data["subtitle"]
    else:
        # Slides standards : on met à jour TITLE, SUBTITLE et BODY
        for shp in iterate_all_shapes(slide.shapes):
            if not shp.has_text_frame:
                continue
            if shp.placeholder_format is not None:
                if shp.placeholder_format.type == PP_PLACEHOLDER.TITLE:
                    shp.text_frame.clear()
                    shp.text = slide_data["title"]
                elif shp.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE:
                    shp.text_frame.clear()
                    shp.text = slide_data["subtitle"]
                elif shp.placeholder_format.type == PP_PLACEHOLDER.BODY:
                    tf = shp.text_frame
                    tf.clear()
                    counters = {}
                    for (block_type, block_para) in slide_data["blocks"]:
                        if block_type == "paragraph":
                            add_paragraph_with_runs(tf, block_para, counters)
            else:
                # Fallback par recherche textuelle
                txt_ref = shp.text.strip().lower()
                if "click to edit master title style" in txt_ref:
                    shp.text_frame.clear()
                    shp.text = slide_data["title"]
                elif "[optional subtitle]" in txt_ref:
                    shp.text_frame.clear()
                    shp.text = slide_data["subtitle"]
                elif "modifiez les styles du texte du masque" in txt_ref:
                    tf = shp.text_frame
                    tf.clear()
                    counters = {}
                    for (block_type, block_para) in slide_data["blocks"]:
                        if block_type == "paragraph":
                            add_paragraph_with_runs(tf, block_para, counters)

# ------------------------------------------------------------------------------------
# Fonction principale de conversion
# ------------------------------------------------------------------------------------
def create_ppt_from_docx(input_docx, template_pptx, output_pptx):
    slides_data = parse_docx_to_slides(input_docx)
    prs = Presentation(template_pptx)

    # Rechercher les layouts par nom
    cover_layout = None
    standard_layout = None
    for layout in prs.slide_layouts:
        if layout.name == "Diapositive de titre":
            cover_layout = layout
        elif layout.name == "Slide_standard layout":
            standard_layout = layout

    if not cover_layout:
        cover_layout = prs.slide_layouts[0]
    if not standard_layout:
        standard_layout = prs.slide_layouts[1]

    # Génération des slides
    for idx, slide_data in enumerate(slides_data):
        layout = cover_layout if idx == 0 else standard_layout
        slide = prs.slides.add_slide(layout)
        clear_shapes_text(slide)
        fill_shapes_text(slide, slide_data, idx)

    prs.save(output_pptx)
    print("Conversion terminée :", output_pptx)

# ------------------------------------------------------------------------------------
# Point d'entrée
# ------------------------------------------------------------------------------------
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage : python convert.py input.docx output.pptx")
        sys.exit(1)
    input_docx = sys.argv[1]
    output_pptx = sys.argv[2]
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_file = os.path.join(script_dir, "template_CVA.pptx")
    create_ppt_from_docx(input_docx, template_file, output_pptx)
