#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script de conversion d'un fichier Word en présentation PowerPoint.

Usage : python convert.py input.docx output.pptx

Le fichier Word doit être structuré ainsi :
  - Chaque slide commence par une ligne "SLIDE X" (ex. : "SLIDE 1")
  - Une ligne "Titre :" définit le titre
  - Une ligne "Sous-titre / Message clé :" définit le sous-titre
  - Le reste du texte constitue le contenu de la slide

La première slide utilisera le layout index 0 du template, les suivantes le layout index 1.
Ce script essaie de préserver le formatage des paragraphes (bullet points et indentations).
"""

import sys
import os
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt

def is_bullet(paragraph):
    """Retourne True si le paragraphe contient une numérotation/bullet (en recherchant <w:numPr> dans son XML)."""
    return "<w:numPr>" in paragraph._p.xml

def get_indentation_level(paragraph):
    """Calcule un niveau d'indentation basé sur la propriété left_indent.
       On considère qu'environ 0.25 inches correspond à 1 niveau d'indentation.
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
    Chaque slide est un dictionnaire contenant :
      - "title" : texte extrait après "Titre :"
      - "subtitle" : texte extrait après "Sous-titre / Message clé :"
      - "blocks" : liste de tuples ("paragraph", paragraph_object)
    """
    doc = Document(doc_path)
    slides = []
    current_slide = None
    for para in doc.paragraphs:
        txt = para.text.strip()
        if not txt:
            continue
        # Nouvelle slide détectée
        if txt.upper().startswith("SLIDE"):
            if current_slide is not None:
                slides.append(current_slide)
            current_slide = {"title": "", "subtitle": "", "blocks": []}
            continue
        # Extraction du titre
        if txt.startswith("Titre :"):
            if current_slide is not None:
                current_slide["title"] = txt[len("Titre :"):].strip()
            continue
        # Extraction du sous-titre
        if txt.startswith("Sous-titre / Message clé :"):
            if current_slide is not None:
                current_slide["subtitle"] = txt[len("Sous-titre / Message clé :"):].strip()
            continue
        # Ajoute le paragraphe complet pour le contenu
        if current_slide is not None:
            current_slide["blocks"].append(("paragraph", para))
    if current_slide is not None:
        slides.append(current_slide)
    return slides

def add_paragraph_with_runs(text_frame, paragraph):
    """
    Ajoute un paragraphe dans le text_frame en copiant les runs du paragraphe Word,
    en préservant le formatage, et en définissant le niveau si le paragraphe est en liste.
    """
    p = text_frame.add_paragraph()
    if is_bullet(paragraph):
        p.level = get_indentation_level(paragraph)
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

def create_ppt_from_docx(doc_path, template_path, output_path):
    """
    Crée une présentation PowerPoint à partir du document Word et d'un template.
    - Utilise le layout index 0 pour la première slide (couverture) et l'index 1 pour les autres.
    - Pour chaque slide, ajoute une zone de texte pour le titre, une pour le sous-titre,
      et une pour le contenu (chaque paragraphe est ajouté en conservant ses bullets et indentations).
    """
    slides_data = parse_docx_to_slides(doc_path)
    prs = Presentation(template_path)
    
    for idx, slide_data in enumerate(slides_data):
        layout = prs.slide_layouts[0] if idx == 0 else prs.slide_layouts[1]
        slide = prs.slides.add_slide(layout)
        
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
        
        # Zone de texte pour le contenu (tout le contenu de la slide)
        content_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(4))
        content_frame = content_box.text_frame
        content_frame.clear()
        for block_type, block_obj in slide_data["blocks"]:
            if block_type == "paragraph":
                add_paragraph_with_runs(content_frame, block_obj)
    
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