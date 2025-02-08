#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Conversion d'un document Word (.docx) en présentation PowerPoint (.pptx)
en utilisant un template existant (template_CVA.pptx).

Le document Word doit respecter la structure suivante :
  • Chaque slide commence par une ligne "SLIDE X"
  • Une ligne "Titre :" suivie du titre
  • Une ligne "Sous-titre / Message clé :" suivie du sous‑titre
  • Le reste (paragraphes, listes, …) constitue le contenu

Le template PowerPoint doit contenir deux layouts :
  • "Diapositive de titre" (pour la slide 0) qui comporte des placeholders
    – Celui dont le texte par défaut contient "PROJECT TITLE" (ou index 11) recevra le titre
    – Celui dont le texte par défaut contient "CVA Presentation title" (ou index 13) recevra le sous‑titre
  • "Slide_standard layout" (pour les slides 1+) qui comporte des placeholders
    – Un placeholder avec le texte "Click to edit Master title style" pour le titre
    – Un placeholder avec le texte "[Optional subtitle]" pour le sous‑titre
    – Un placeholder avec le texte "Modifiez les styles du texte du masque" pour le contenu,
      auquel on injecte le texte du Word en conservant le formatage.

Usage :
  python convert.py input.docx output.pptx

Le fichier template_CVA.pptx doit se trouver dans le même dossier que ce script.
"""

import sys
import os
from docx import Document
from pptx import Presentation
from pptx.util import Pt

# -----------------------------
# Extraction du contenu du Word
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
    # Debug: Afficher le contenu extrait
    print("Contenu extrait du Word:")
    for i, slide in enumerate(slides_data):
        print(f"Slide {i} - Titre: {slide['title']} | Sous-titre: {slide['subtitle']} | Nb de blocs: {len(slide['blocks'])}")
    return slides_data

# -----------------------------
# Gestion des listes et mise en forme
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
# Remplissage des placeholders pour la slide de couverture
# -----------------------------
def fill_cover_slide(slide, slide_data):
    print("Placeholders dans la slide de couverture:")
    for ph in slide.placeholders:
        try:
            idx = ph.placeholder_format.idx
        except Exception:
            idx = "n/a"
        print(f" - Index {idx}: '{ph.text}'")
    try:
        # Selon votre debug, on souhaite que :
        # Le placeholder d'index 11 contienne le titre
        slide.placeholders[11].text = slide_data["title"]
        print("Placeholder 11 mis à jour avec le titre:", slide_data["title"])
    except Exception as e:
        print("Erreur lors de l'insertion dans le placeholder index 11:", e)
    try:
        # Le placeholder d'index 13 contiendra le sous-titre
        slide.placeholders[13].text = slide_data["subtitle"]
        print("Placeholder 13 mis à jour avec le sous-titre:", slide_data["subtitle"])
    except Exception as e:
        print("Erreur lors de l'insertion dans le placeholder index 13:", e)

# -----------------------------
# Remplissage des placeholders pour les slides standards
# -----------------------------
def fill_standard_slide(slide, slide_data):
    print("Placeholders dans une slide standard:")
    for ph in slide.placeholders:
        try:
            idx = ph.placeholder_format.idx
        except Exception:
            idx = "n/a"
        print(f" - Index {idx}: '{ph.text}'")
    try:
        # Pour le titre : placeholder index 3 doit recevoir le titre
        slide.placeholders[3].text = slide_data["title"]
        print("Placeholder 3 mis à jour avec le titre:", slide_data["title"])
    except Exception as e:
        print("Erreur pour le titre (index 3):", e)
    try:
        # Pour le sous-titre : placeholder index 2 doit recevoir le sous-titre
        slide.placeholders[2].text = slide_data["subtitle"]
        print("Placeholder 2 mis à jour avec le sous-titre:", slide_data["subtitle"])
    except Exception as e:
        print("Erreur pour le sous-titre (index 2):", e)
    try:
        # Pour le contenu : placeholder index 7
        ph_content = slide.placeholders[7]
        ph_content.text_frame.clear()
        counters = {}
        for (block_type, block_para) in slide_data["blocks"]:
            if block_type == "paragraph":
                add_paragraph_with_runs(ph_content.text_frame, block_para, counters)
        print("Placeholder 7 rempli avec le contenu.")
    except Exception as e:
        print("Erreur pour le contenu (index 7):", e)

# -----------------------------
# Fonction principale de conversion
# -----------------------------
def create_ppt_from_docx(input_docx, template_pptx, output_pptx):
    slides_data = parse_docx_to_slides(input_docx)
    prs = Presentation(template_pptx)

    # Recherche des layouts par nom
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
        print("=== Traitement de la slide 0 ===")
        fill_cover_slide(slide0, slides_data[0])
    else:
        print("Aucune slide trouvée dans le document Word.")

    # Créer les slides standards (slides 1 à N)
    for idx, slide_data in enumerate(slides_data[1:], start=1):
        slide = prs.slides.add_slide(standard_layout)
        print(f"=== Traitement de la slide {idx} ===")
        fill_standard_slide(slide, slide_data)

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
