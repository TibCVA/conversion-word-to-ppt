#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script de conversion d'un fichier Word en présentation PowerPoint en insérant
le contenu dans les placeholders existants du template.

Usage : python convert.py input.docx output.pptx

Structure attendue du fichier Word (input.docx) :
  • Chaque slide commence par une ligne "SLIDE X" (ex. "SLIDE 1")
  • Une ligne "Titre :" définit le titre
  • Une ligne "Sous-titre / Message clé :" définit le sous-titre
  • Le reste du texte constitue le contenu de la slide (paragraphes, listes, etc.)

Le template contient :
  • Pour SLIDE 0 (couverture) : les placeholders "PROJECT TITLE", "CVA Presentation title", "Subtitle"
  • Pour SLIDES 1 à N (standard) : les placeholders dont le texte par défaut est
       "Click to edit Master title style" (titre),
       "[Optional subtitle]" (sous-titre) et
       "Modifiez les styles du texte du masque…" (body)

Ce script insère le contenu dans ces zones en essayant de préserver la mise en forme (puces, numérotation, gras, souligné, etc.).
"""

import sys
import os
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_PLACEHOLDER  # Utilisation de MSO_PLACEHOLDER

def get_list_type(paragraph):
    """
    Retourne "bullet" si le paragraphe utilise une puce (w:numFmt val="bullet"),
    "decimal" s'il est numéroté (w:numFmt val="decimal"), ou None sinon.
    """
    xml = paragraph._p.xml
    if 'w:numFmt val="bullet"' in xml:
        return "bullet"
    elif 'w:numFmt val="decimal"' in xml:
        return "decimal"
    else:
        return None

def get_indentation_level(paragraph):
    """
    Estime le niveau d'indentation à partir de paragraph_format.left_indent.
    On considère environ 0.25 inch par niveau.
    """
    indent = paragraph.paragraph_format.left_indent
    if indent is None:
        return 0
    try:
        return int(indent.inches / 0.25)
    except Exception:
        return 0

def parse_docx_to_slides(doc_path):
    """
    Parcourt le document Word et découpe son contenu en slides.
    Retourne une liste de dictionnaires, chacun contenant :
      - "title": le texte après "Titre :"
      - "subtitle": le texte après "Sous-titre / Message clé :"
      - "blocks": liste d'objets Paragraph (le contenu de la slide)
    """
    doc = Document(doc_path)
    slides = []
    current_slide = None
    for para in doc.paragraphs:
        txt = para.text.strip()
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
            current_slide["blocks"].append(para)
    if current_slide is not None:
        slides.append(current_slide)
    return slides

def add_paragraph_with_runs(text_frame, paragraph, counters):
    """
    Ajoute un paragraphe dans le text_frame en copiant ses runs pour préserver le formatage.
    Si le paragraphe est en liste, préfixe avec :
      - "• " pour une liste à puces,
      - ou avec un numéro (calculé par niveau) pour une liste numérotée.
    'counters' est un dictionnaire qui maintient le compte par niveau.
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
            r.font.size = run.font.size if run.font.size else Pt(14)
        if paragraph.alignment is not None:
            p.alignment = paragraph.alignment
        return p

def clear_all_placeholders(slide):
    """
    Efface le texte de tous les placeholders de la diapositive afin d'éviter d'afficher le texte par défaut.
    """
    for shape in slide.placeholders:
        try:
            shape.text = ""
        except Exception:
            pass

def fill_placeholders(slide, slide_data, slide_index):
    """
    Remplit les placeholders de la diapositive avec le contenu extrait du Word.
    Pour SLIDE 0, on remplace les placeholders "PROJECT TITLE", "CVA Presentation title" et "Subtitle".
    Pour SLIDES 1+, on remplace :
      - le placeholder de titre (contenant "Click to edit Master title style") par le titre,
      - le placeholder de sous-titre (contenant "[Optional subtitle]") par le sous-titre,
      - le placeholder de contenu (celui dont le texte par défaut débute par "Modifiez les styles du texte du masque") par le contenu.
    Les autres placeholders (notes, sources, numéros) restent intacts.
    """
    if slide_index == 0:
        for shape in slide.placeholders:
            if not shape.has_text_frame:
                continue
            txt = shape.text.strip()
            if txt == "PROJECT TITLE":
                shape.text = slide_data["title"]
            elif txt == "CVA Presentation title":
                shape.text = slide_data["title"]
            elif txt == "Subtitle":
                shape.text = slide_data["subtitle"]
    else:
        for shape in slide.placeholders:
            if not shape.has_text_frame:
                continue
            txt = shape.text.strip()
            if txt.startswith("Click to edit Master title style"):
                shape.text = slide_data["title"]
            elif txt.startswith("[Optional subtitle]"):
                shape.text = slide_data["subtitle"]
            elif txt.startswith("Modifiez les styles du texte du masque"):
                tf = shape.text_frame
                tf.clear()
                counters = {}
                for para in slide_data["blocks"]:
                    add_paragraph_with_runs(tf, para, counters)

def create_ppt_from_docx(doc_path, template_path, output_path):
    """
    Crée une présentation PowerPoint à partir du document Word et du template,
    en insérant le contenu dans les placeholders existants.
    """
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
