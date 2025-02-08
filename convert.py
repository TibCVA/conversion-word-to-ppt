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
"""

import sys
import os
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt

def parse_docx_to_slides(doc_path):
    """
    Lit le fichier Word et découpe son contenu en slides.
    Chaque slide est représentée par un dictionnaire contenant :
      - "title" : texte extrait après "Titre :"
      - "subtitle" : texte extrait après "Sous-titre / Message clé :"
      - "blocks" : liste de chaînes de caractères (les paragraphes du contenu)
    """
    doc = Document(doc_path)
    slides = []
    current_slide = None
    for para in doc.paragraphs:
        txt = para.text.strip()
        if not txt:
            continue
        # Début d'une nouvelle slide
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
        # Extraction du sous-titre / message clé
        if txt.startswith("Sous-titre / Message clé :"):
            if current_slide is not None:
                current_slide["subtitle"] = txt[len("Sous-titre / Message clé :"):].strip()
            continue
        # Autres paragraphes : ajout au contenu
        if current_slide is not None:
            current_slide["blocks"].append(txt)
    if current_slide is not None:
        slides.append(current_slide)
    return slides

def create_ppt_from_docx(doc_path, template_path, output_path):
    """
    Crée une présentation PowerPoint à partir du document Word et du template.
    - Pour la première slide, utilise le layout index 0.
    - Pour les slides suivantes, utilise le layout index 1.
    Pour chaque slide, le script ajoute :
      • Une zone de texte pour le titre
      • Une zone de texte pour le sous-titre
      • Une zone de texte pour le contenu (chaque paragraphe est ajouté comme un paragraphe séparé)
    """
    slides_data = parse_docx_to_slides(doc_path)
    prs = Presentation(template_path)
    
    for idx, slide_data in enumerate(slides_data):
        # Choix du layout : index 0 pour la première slide, index 1 pour les autres.
        layout = prs.slide_layouts[0] if idx == 0 else prs.slide_layouts[1]
        slide = prs.slides.add_slide(layout)
        
        # Ajout du titre (zone de texte positionnée en haut)
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
        title_frame = title_box.text_frame
        title_frame.clear()
        title_frame.text = slide_data["title"]
        
        # Ajout du sous-titre (zone de texte juste en dessous du titre)
        subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(1))
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.clear()
        subtitle_frame.text = slide_data["subtitle"]
        
        # Ajout du contenu (zone de texte plus large)
        content_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(4))
        content_frame = content_box.text_frame
        content_frame.clear()
        for block in slide_data["blocks"]:
            p = content_frame.add_paragraph()
            p.text = block

    prs.save(output_path)
    print("Conversion terminée :", output_path)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python convert.py input.docx output.pptx")
        sys.exit(1)
    input_docx = sys.argv[1]
    output_pptx = sys.argv[2]
    # Le fichier template_CVA.pptx doit être présent à la racine du dépôt.
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template_CVA.pptx")
    create_ppt_from_docx(input_docx, template_path, output_pptx)