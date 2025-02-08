#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Conversion d'un document Word (.docx) en présentation PowerPoint (.pptx)
en utilisant un template existant (template_CVA.pptx).

Structure attendue du Word :
  • Chaque slide commence par "SLIDE X"
  • "Titre :" indique le titre
  • "Sous-titre / Message clé :" indique le sous‑titre
  • Le reste (paragraphes, listes, …) constitue le contenu

Mapping des placeholders (indices fixes d'après votre debug) :

LAYOUT #0: Diapositive de titre
  - Placeholder index 11 (type PLACEHOLDER) doit recevoir le texte du champ "Titre :"
  - Placeholder index 13 (type PLACEHOLDER) doit recevoir le texte du champ "Sous-titre / Message clé :"
  - Les autres placeholders ne sont pas modifiés.

LAYOUT #1: Slide_standard layout
  - Placeholder index 3 (type PLACEHOLDER) doit recevoir le texte du champ "Titre :"
  - Placeholder index 2 (type PLACEHOLDER) doit recevoir le texte du champ "Sous-titre / Message clé :"
  - Placeholder index 7 (type PLACEHOLDER) doit être vidé puis rempli avec le contenu textuel du Word,
    en conservant le formatage (bullet points, indentations, etc.).

Usage :
  python convert.py input.docx output.pptx

Le fichier template_CVA.pptx doit être dans le même dossier que ce script.
"""

import sys
import os
from docx import Document
from pptx import Presentation
from pptx.util import Pt

# -----------------------------
# Extraction du contenu Word
# -----------------------------
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

# -----------------------------
# Gestion du formatage (listes, runs)
# -----------------------------
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

# -----------------------------
# Remplissage des placeholders par index avec debug
# -----------------------------
def fill_placeholders_by_index(slide, slide_data, slide_index):
    placeholders = list(slide.placeholders)
    print("Nombre de placeholders dans la slide:", len(placeholders))
    for i, ph in enumerate(placeholders):
        try:
            idx = ph.placeholder_format.idx
        except Exception:
            idx = "n/a"
        print("Placeholder index:", idx, "Texte:", repr(ph.text))
        
    if slide_index == 0:
        print("Remplissage de la slide de couverture (slide 0)")
        try:
            slide.placeholders[11].text = slide_data["title"]
            print(" - Placeholder index 11 mis à jour avec:", slide_data["title"])
        except Exception as e:
            print("Erreur pour placeholder index 11:", e)
        try:
            slide.placeholders[13].text = slide_data["subtitle"]
            print(" - Placeholder index 13 mis à jour avec:", slide_data["subtitle"])
        except Exception as e:
            print("Erreur pour placeholder index 13:", e)
    else:
        print("Remplissage d'une slide standard (slide >=1)")
        try:
            slide.placeholders[3].text = slide_data["title"]
            print(" - Placeholder index 3 mis à jour avec:", slide_data["title"])
        except Exception as e:
            print("Erreur pour placeholder index 3:", e)
        try:
            slide.placeholders[2].text = slide_data["subtitle"]
            print(" - Placeholder index 2 mis à jour avec:", slide_data["subtitle"])
        except Exception as e:
            print("Erreur pour placeholder index 2:", e)
        try:
            ph_body = slide.placeholders[7]
            ph_body.text_frame.clear()
            counters = {}
            for (block_type, block_para) in slide_data["blocks"]:
                if block_type == "paragraph":
                    add_paragraph_with_runs(ph_body.text_frame, block_para, counters)
            print(" - Placeholder index 7 rempli avec le contenu.")
        except Exception as e:
            print("Erreur pour placeholder index 7:", e)

# -----------------------------
# Fonction principale de conversion
# -----------------------------
def create_ppt_from_docx(input_docx, template_pptx, output_pptx):
    slides_data = parse_docx_to_slides(input_docx)
    prs = Presentation(template_pptx)

    # Sélectionner les layouts par nom (ou utiliser des fallback)
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

    # Créer les slides standards (slides 1 à N)
    for idx, slide_data in enumerate(slides_data[1:], start=1):
        slide = prs.slides.add_slide(standard_layout)
        fill_placeholders_by_index(slide, slide_data, idx)

    prs.save(output_pptx)
    print("Conversion terminée :", output_pptx)

# -----------------------------
# Point d'entrée
# -----------------------------
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage : python convert.py input.docx output.pptx")
        sys.exit(1)
    input_docx = sys.argv[1]
    output_pptx = sys.argv[2]
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_file = os.path.join(script_dir, "template_CVA.pptx")
    create_ppt_from_docx(input_docx, template_file, output_pptx)
