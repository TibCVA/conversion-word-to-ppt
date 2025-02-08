#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Conversion d'un document Word (.docx) en présentation PowerPoint (.pptx)
en utilisant un template existant (template_CVA.pptx).

Structure du Word attendue :
  • Chaque slide commence par "SLIDE X"
  • "Titre :" indique le titre
  • "Sous-titre / Message clé :" indique le sous-titre
  • Le reste (paragraphes, listes, …) constitue le contenu

Ce script utilise les indices des placeholders (fixés d'après le debug) pour insérer :
  - Pour la slide 0 (couverture) :
       • slide.placeholders[11] → titre
       • slide.placeholders[13] → sous‑titre
  - Pour les slides standards (slides 1+) :
       • slide.placeholders[3] → titre
       • slide.placeholders[2] → sous‑titre
       • slide.placeholders[7] → contenu (BODY), auquel on insère le texte du Word

Usage :
  python convert.py input.docx output.pptx

Le fichier template_CVA.pptx doit se trouver dans le même dossier que ce script.
"""

import sys
import os
from docx import Document
from pptx import Presentation
from pptx.util import Pt

# --------------------------------------------------------------------
# Extraction du contenu du Word en une liste de slides
# --------------------------------------------------------------------
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

# --------------------------------------------------------------------
# Gestion des listes et insertion de runs pour conserver le formatage
# --------------------------------------------------------------------
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

# --------------------------------------------------------------------
# Remplissage des placeholders par index
# --------------------------------------------------------------------
def fill_placeholders_by_index(slide, slide_data, slide_index):
    placeholders = list(slide.placeholders)
    if slide_index == 0:
        # Slide de couverture
        try:
            # On force l'insertion dans les placeholders d'index 11 et 13
            placeholders[11].text = slide_data["title"]
            placeholders[13].text = slide_data["subtitle"]
        except IndexError:
            print("Erreur : nombre insuffisant de placeholders pour la slide de couverture.")
    else:
        try:
            placeholders[3].text = slide_data["title"]
            placeholders[2].text = slide_data["subtitle"]
            # Pour le contenu, on utilise le placeholder index 7 (BODY)
            body_ph = placeholders[7]
            body_tf = body_ph.text_frame
            body_tf.clear()
            counters = {}
            for (block_type, block_para) in slide_data["blocks"]:
                if block_type == "paragraph":
                    add_paragraph_with_runs(body_tf, block_para, counters)
        except IndexError:
            print("Erreur : nombre insuffisant de placeholders pour une slide standard.")

# --------------------------------------------------------------------
# Fonction principale de conversion
# --------------------------------------------------------------------
def create_ppt_from_docx(input_docx, template_pptx, output_pptx):
    slides_data = parse_docx_to_slides(input_docx)
    prs = Presentation(template_pptx)
    
    # Pour la slide de couverture, on utilise le layout "Diapositive de titre"
    # Pour les autres, on utilise le layout "Slide_standard layout"
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

    # Créer la slide de couverture (slide 0)
    if slides_data:
        slide0 = prs.slides.add_slide(cover_layout)
        fill_placeholders_by_index(slide0, slides_data[0], 0)
    else:
        print("Aucune diapositive trouvée dans le document Word.")
    
    # Créer les slides standards (slides 1+)
    for idx, slide_data in enumerate(slides_data[1:], start=1):
        slide = prs.slides.add_slide(standard_layout)
        fill_placeholders_by_index(slide, slide_data, idx)
    
    prs.save(output_pptx)
    print("Conversion terminée :", output_pptx)

# --------------------------------------------------------------------
# Point d'entrée
# --------------------------------------------------------------------
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage : python convert.py input.docx output.pptx")
        sys.exit(1)
    input_docx = sys.argv[1]
    output_pptx = sys.argv[2]
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_file = os.path.join(script_dir, "template_CVA.pptx")
    create_ppt_from_docx(input_docx, template_file, output_pptx)
