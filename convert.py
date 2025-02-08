#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script de conversion d'un fichier Word en présentation PowerPoint.

Usage : python convert.py input.docx output.pptx

Structure attendue du fichier Word (input.docx) :
  - Chaque slide commence par une ligne "SLIDE X" (ex. : "SLIDE 1")
  - Une ligne "Titre :" définit le titre
  - Une ligne "Sous-titre / Message clé :" définit le sous-titre
  - Le reste du texte constitue le contenu de la slide,
    qui peut contenir des paragraphes en liste (à puces ou numérotés) et des tableaux.

La première slide utilisera le layout index 0 du template (template_CVA.pptx),
les suivantes le layout index 1.

Ce script tente de :
  • Regrouper tout le contenu d'une slide sur une seule diapositive.
  • Reproduire le type de liste :  
      - Si le paragraphe est une liste à puces (w:numFmt val="bullet"), préfixer avec "• ".
      - Si c'est une liste numérotée (w:numFmt val="decimal"), préfixer avec un numéro basé sur le niveau.
  • Conserver le formatage des runs (gras, souligné, taille, etc.).
  • Copier les tableaux en recréant la structure.
  • Effacer complètement les placeholders par défaut du master.
"""

import sys
import os
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt

def get_list_type(paragraph):
    """
    Retourne "bullet" si le paragraphe utilise une puce,
    "decimal" s'il est numéroté, ou None sinon.
    Pour cela, on inspecte le XML du paragraphe.
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
    Estime le niveau d'indentation basé sur left_indent.
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
    Lit le document Word et découpe son contenu en slides.
    Retourne une liste de dictionnaires, chacun contenant :
      - "title" : texte après "Titre :"
      - "subtitle" : texte après "Sous-titre / Message clé :"
      - "blocks" : liste d'objets Paragraph (pour le texte) ou Table (pour les tableaux)
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
    Ajoute un paragraphe dans le text_frame en copiant ses runs.
    Si le paragraphe est en liste, préfixe avec :
      - "• " pour une liste à puces,
      - ou avec un numéro (basé sur le niveau) pour une liste numérotée.
    La variable counters est un dictionnaire qui conserve le compte par niveau.
    """
    p = text_frame.add_paragraph()
    list_type = get_list_type(paragraph)
    level = get_indentation_level(paragraph)
    if list_type == "bullet":
        # Conserver le bullet
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
        # Copie simple des runs pour respecter le formatage
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

def add_table(slide, table, left, top, width, height):
    """
    Ajoute un tableau sur la diapositive en recréant la structure.
    Pour chaque cellule, concatène les paragraphes (avec retour à la ligne).
    Note : La reproduction exacte du formatage de tableau est limitée.
    """
    rows = len(table.rows)
    cols = len(table.columns)
    shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    ppt_table = shape.table
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            cell_text = "\n".join(p.text for p in cell.paragraphs)
            ppt_table.cell(i, j).text = cell_text
    return shape

def clear_placeholders(slide):
    """
    Supprime ou efface les zones de texte des placeholders du master afin qu'ils ne soient pas visibles.
    Ici, nous parcourons les formes et si elles sont des placeholders, on les masque.
    """
    for shape in slide.shapes:
        if hasattr(shape, "is_placeholder") and shape.is_placeholder:
            try:
                shape.text = ""
            except Exception:
                pass

def create_ppt_from_docx(doc_path, template_path, output_path):
    """
    Crée une présentation PowerPoint à partir du document Word et d'un template.
    Pour chaque slide, ajoute :
      - Une zone de texte pour le titre,
      - Une zone de texte pour le sous-titre,
      - Une zone de texte pour le contenu.
    Le contenu des listes est traité pour différencier puces et numérotation.
    """
    slides_data = parse_docx_to_slides(doc_path)
    prs = Presentation(template_path)
    
    for idx, slide_data in enumerate(slides_data):
        layout = prs.slide_layouts[0] if idx == 0 else prs.slide_layouts[1]
        slide = prs.slides.add_slide(layout)
        
        # Effacer les placeholders existants
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
        # Un dictionnaire pour gérer la numérotation par niveau
        numbered_counters = {}
        for paragraph in slide_data["blocks"]:
            add_paragraph_with_runs(content_frame, paragraph, numbered_counters)
    
    # Pour les tableaux, on peut parcourir le document Word et ajouter les tableaux si besoin.
    # Si votre fichier Word comporte des tableaux insérés en dehors des paragraphes,
    # il faudra étendre cette logique.
    
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
