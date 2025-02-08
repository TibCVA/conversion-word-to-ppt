#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Conversion d'un document Word (.docx) en présentation PowerPoint (.pptx)
en utilisant un template existant (template_CVA.pptx).

Structure du Word attendue :
  • Chaque slide commence par "SLIDE X"
  • "Titre :" suivi du titre
  • "Sous-titre / Message clé :" suivi du sous-titre
  • Le reste (paragraphes, listes, etc.) constitue le contenu

Le template PPT doit contenir deux layouts :
  – "Diapositive de titre" (pour la slide 0) avec :
       • Un placeholder dont le texte contient "PROJECT TITLE" → recevra le titre
       • Un placeholder dont le texte contient "CVA Presentation title" → recevra le sous‑titre
       • Un placeholder dont le texte contient "Subtitle" → sera vidé
  – "Slide_standard layout" (pour les slides 1+) avec :
       • Un placeholder contenant "Click to edit Master title style" → recevra le titre
       • Un placeholder contenant "[Optional subtitle]" → recevra le sous‑titre
       • Un placeholder contenant "Modifiez les styles du texte du masque" → sera vidé et rempli avec le contenu du Word
       
Usage :
  python convert.py input.docx output.pptx

Le template (template_CVA.pptx) doit se trouver dans le même dossier que ce script.
"""

import sys
import os
from docx import Document
from pptx import Presentation
from pptx.util import Pt

# ------------------------------------------------------------------------------------
# Extraction du contenu du Word en une liste de slides
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
            # On enregistre le paragraphe et sa mise en forme (pour traitement ultérieur)
            current_slide["blocks"].append(("paragraph", para))

    if current_slide is not None:
        slides_data.append(current_slide)
    return slides_data

# ------------------------------------------------------------------------------------
# Gestion des listes et formatage (bullet, indentations, etc.)
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
# Remplissage des placeholders de la slide en se basant sur slide.placeholders
# ------------------------------------------------------------------------------------
def fill_placeholders(slide, slide_data, slide_index):
    # On parcourt uniquement les placeholders de la slide
    for ph in slide.placeholders:
        txt = ph.text.strip().lower() if ph.text else ""
        if slide_index == 0:
            # Pour la slide de couverture (layout "Diapositive de titre")
            if "project title" in txt:
                ph.text = slide_data["title"]
            elif "cva presentation title" in txt:
                ph.text = slide_data["subtitle"]
            elif "subtitle" in txt:
                ph.text = ""  # On vide ce placeholder
        else:
            # Pour les slides standards (layout "Slide_standard layout")
            if "[optional subtitle]" in txt:
                ph.text = slide_data["subtitle"]
            elif "click to edit master title style" in txt:
                ph.text = slide_data["title"]
            elif "modifiez les styles du texte du masque" in txt:
                # On vide puis on insère le contenu du Word (en conservant le formatage)
                tf = ph.text_frame
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

    # Fallback si non trouvés
    if not cover_layout:
        cover_layout = prs.slide_layouts[0]
    if not standard_layout:
        standard_layout = prs.slide_layouts[1]

    # Pour chaque slide du Word, on crée une diapositive dans le PPT
    for idx, slide_data in enumerate(slides_data):
        layout = cover_layout if idx == 0 else standard_layout
        slide = prs.slides.add_slide(layout)
        # Remplir les placeholders
        fill_placeholders(slide, slide_data, idx)

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
