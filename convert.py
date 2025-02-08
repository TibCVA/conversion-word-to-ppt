#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script de conversion d'un fichier Word en présentation PowerPoint.

Usage : python convert.py input.docx output.pptx

Structure attendue du fichier Word (input.docx) :
  - Chaque slide commence par une ligne "SLIDE X" (ex. : "SLIDE 1")
  - Une ligne "Titre :" définit le titre
  - Une ligne "Sous-titre / Message clé :" définit le sous-titre
  - Le reste du texte constitue le contenu de la slide

La première slide utilisera le layout index 0 du template (template_CVA.pptx),
les suivantes utiliseront le layout index 1.

Le script tente de :
  • Garder tout le contenu d'une slide sur la même diapositive.
  • Reproduire les bullet points et la numérotation.
  • Conserver le formatage des runs (gras, souligné, etc.).
  • Effacer le texte par défaut issu des placeholders du master.
"""

import sys
import os
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt

def is_bullet(paragraph):
    """
    Retourne True si le paragraphe semble être en liste (détecte la présence
    du tag <w:numPr> ou si le style du paragraphe contient "List" ou "Bullet").
    """
    style_name = paragraph.style.name if paragraph.style else ""
    if "bullet" in style_name.lower() or "list" in style_name.lower():
        return True
    return "<w:numPr>" in paragraph._p.xml

def get_indentation_level(paragraph):
    """
    Estime le niveau d'indentation à partir de paragraph_format.left_indent.
    On considère qu'environ 0.25 inch correspond à 1 niveau.
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
    Lit le document Word et découpe son contenu en slides.
    Retourne une liste de dictionnaires, chacun ayant :
      - "title" : texte après "Titre :"
      - "subtitle" : texte après "Sous-titre / Message clé :"
      - "blocks" : liste d'objets Paragraph (tout le contenu de la slide)
    """
    doc = Document(doc_path)
    slides = []
    current_slide = None
    for para in doc.paragraphs:
        txt = para.text.strip()
        if not txt:
            continue
        # Détection d'un marqueur de nouvelle slide
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
    Ajoute un paragraphe dans le text_frame en copiant ses runs pour conserver le formatage.
    Si le paragraphe est en liste, préfixe avec :
      - "• " pour une liste à puces (si le style contient "bullet"),
      - ou avec un numéro pour une liste numérotée, en se basant sur le niveau d'indentation.
    La variable counters (dictionnaire) maintient le compte par niveau d'indentation.
    """
    p = text_frame.add_paragraph()
    style_name = paragraph.style.name if paragraph.style else ""
    if is_bullet(paragraph):
        # Liste à puces ou numérotée
        if "bullet" in style_name.lower():
            prefix = "• "
        else:
            level = get_indentation_level(paragraph)
            count = counters.get(level, 0) + 1
            counters[level] = count
            prefix = f"{count}. "
        # On ajoute le préfixe suivi du texte du paragraphe
        p.text = prefix + paragraph.text
        p.level = get_indentation_level(paragraph)
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

def clear_placeholders(slide):
    """Efface le texte des placeholders par défaut sur la diapositive."""
    for shape in slide.shapes:
        if hasattr(shape, "is_placeholder") and shape.is_placeholder:
            try:
                shape.text = ""
            except Exception:
                pass

def create_ppt_from_docx(doc_path, template_path, output_path):
    """
    Crée une présentation PowerPoint à partir d'un document Word et d'un template.
    Pour chaque slide, ajoute :
      - Une zone de texte pour le titre (position fixe)
      - Une zone de texte pour le sous-titre (position fixe)
      - Une zone de texte pour le contenu (position fixe)
    Tout le contenu d'une slide est regroupé sur une seule diapositive.
    """
    slides_data = parse_docx_to_slides(doc_path)
    prs = Presentation(template_path)
    
    for idx, slide_data in enumerate(slides_data):
        layout = prs.slide_layouts[0] if idx == 0 else prs.slide_layouts[1]
        slide = prs.slides.add_slide(layout)
        
        # Effacer les placeholders par défaut
        clear_placeholders(slide)
        
        # Zone de texte pour le titre
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
        title_frame = title_box.text_frame
        title_frame.clear()
        title_frame.text = slide_data["title"]
        
        # Zone de texte pour le sous-titre
        subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(1))
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.clear()
        subtitle_frame.text = slide_data["subtitle"]
        
        # Zone de texte pour le contenu
        content_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(4))
        content_frame = content_box.text_frame
        content_frame.clear()
        # Dictionnaire pour suivre la numérotation par niveau dans la slide
        numbered_counters = {}
        for paragraph in slide_data["blocks"]:
            add_paragraph_with_runs(content_frame, paragraph, numbered_counters)
    
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