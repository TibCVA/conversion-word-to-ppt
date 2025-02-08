#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Conversion d'un document Word en présentation PowerPoint en insérant
le contenu dans les placeholders existants du template.

Usage : python convert.py input.docx output.pptx

Structure attendue du Word (input.docx) :
  • Chaque slide commence par une ligne "SLIDE X" (ex. "SLIDE 1")
  • Une ligne "Titre :" définit le titre de la slide
  • Une ligne "Sous-titre / Message clé :" définit le sous-titre
  • Le reste (paragraphes et tableaux) constitue le contenu

Template PPT :
  • Pour SLIDE 0 (couverture) : placeholders dont le texte par défaut contient 
      "PROJECT TITLE", "CVA Presentation title" et "Subtitle" (le placeholder "Date" reste inchangé)
  • Pour SLIDES 1+ (standard) : placeholders dont le texte par défaut contient 
      "Click to edit Master title style" (titre), "[Optional subtitle]" (sous-titre) et 
      "Modifiez les styles du texte du masque" (body)

Le script insère le contenu dans ces zones en préservant autant que possible :
  - Les listes à puces ou numérotées (en utilisant le style et le niveau d'indentation)
  - Le formatage des runs (gras, italique, souligné, taille)
  - Les tableaux du Word en tant que tableaux modifiables
"""

import sys
import os
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt

# --- Fonctions de détection et de traitement des listes ---

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
    Environ 0.25 inch correspond à 1 niveau.
    """
    indent = paragraph.paragraph_format.left_indent
    if indent is None:
        return 0
    try:
        return int(indent.inches / 0.25)
    except Exception:
        return 0

# --- Extraction du contenu du Word en slides ---
def parse_docx_to_slides(doc_path):
    """
    Parcourt le document Word et découpe son contenu en slides.
    Retourne une liste de dictionnaires avec :
      - "title": le texte après "Titre :"
      - "subtitle": le texte après "Sous-titre / Message clé :"
      - "blocks": liste d'objets (tuples ("paragraph", paragraph) ou ("table", table))
    """
    doc = Document(doc_path)
    slides = []
    current_slide = None
    # On parcourt l'ensemble des blocs (paragraphes et tableaux) dans l'ordre du document.
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
            current_slide["blocks"].append(("paragraph", para))
    # Traitement des tableaux (ajouter tous les tableaux trouvés, s'ils existent)
    for table in doc.tables:
        # On ajoute le tableau dans la dernière slide (si nécessaire, vous pouvez adapter cette logique)
        if current_slide is not None:
            current_slide["blocks"].append(("table", table))
    if current_slide is not None:
        slides.append(current_slide)
    return slides

# --- Insertion du texte dans un text frame en préservant le formatage ---
def add_paragraph_with_runs(text_frame, paragraph, counters):
    """
    Ajoute un paragraphe dans le text_frame en copiant ses runs.
    Si le paragraphe est une liste, préfixe avec "• " pour les puces ou avec un numéro pour les listes numérotées.
    'counters' est un dictionnaire qui garde la numérotation par niveau.
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
        return p

# --- Insertion d'un tableau Word dans PowerPoint ---
def insert_table(slide, table, left, top, width, height):
    """
    Insère un tableau modifiable dans la diapositive.
    Respecte le nombre de lignes et de colonnes du tableau Word.
    Les cellules fusionnées ne sont pas automatiquement détectées ici (cette opération est complexe),
    mais la structure de base et le contenu textuel de chaque cellule sont copiés.
    """
    rows = len(table.rows)
    cols = len(table.columns)
    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    ppt_table = table_shape.table
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            cell_text = "\n".join(p.text for p in cell.paragraphs)
            ppt_table.cell(i, j).text = cell_text
    return table_shape

# --- Remplissage des placeholders du template ---
def fill_placeholders(slide, slide_data, slide_index):
    """
    Remplit les zones réservées (placeholders) de la diapositive avec le contenu extrait.
    Pour SLIDE 0, on remplace les placeholders dont le texte par défaut contient :
         "PROJECT TITLE", "CVA Presentation title", "Subtitle"
    Pour les SLIDES 1+, on remplace les placeholders dont le texte par défaut contient :
         "Click to edit Master title style" (titre),
         "[Optional subtitle]" (sous-titre) et
         "Modifiez les styles du texte du masque" (body).
    Les autres placeholders (notes, sources, numéros) restent inchangés.
    """
    if slide_index == 0:
        for shape in slide.placeholders:
            if not shape.has_text_frame:
                continue
            txt = shape.text.strip().lower()
            if "project title" in txt:
                shape.text = slide_data["title"]
            elif "cva presentation title" in txt:
                shape.text = slide_data["title"]
            elif "subtitle" in txt:
                shape.text = slide_data["subtitle"]
    else:
        for shape in slide.placeholders:
            if not shape.has_text_frame:
                continue
            txt = shape.text.strip().lower()
            if "click to edit master title style" in txt:
                shape.text = slide_data["title"]
            elif "[optional subtitle]" in txt:
                shape.text = slide_data["subtitle"]
            elif "modifiez les styles du texte du masque" in txt:
                tf = shape.text_frame
                tf.clear()
                counters = {}
                for block_type, block in slide_data["blocks"]:
                    if block_type == "paragraph":
                        add_paragraph_with_runs(tf, block, counters)
                    elif block_type == "table":
                        # Si le contenu contient des tableaux, insérez-les après le texte dans ce placeholder.
                        # Ici, nous ajoutons un paragraphe vide pour séparer le texte du tableau.
                        tf.add_paragraph()
                        # Utilisez les coordonnées du placeholder pour insérer le tableau.
                        left = shape.left
                        top = shape.top + shape.height + Inches(0.2)
                        width = shape.width
                        # Estimez la hauteur du tableau en fonction du nombre de lignes (0.8 inch par ligne par exemple)
                        height = Inches(0.8 * len(block.rows))
                        insert_table(slide, block, left, top, width, height)

def clear_all_placeholders(slide):
    """
    Efface le texte de tous les placeholders de la diapositive pour éviter l'affichage du texte par défaut.
    """
    for shape in slide.placeholders:
        try:
            shape.text = ""
        except Exception:
            pass

# --- Fonction principale de conversion ---
def create_ppt_from_docx(doc_path, template_path, output_path):
    slides_data = parse_docx_to_slides(doc_path)
    prs = Presentation(template_path)
    for idx, slide_data in enumerate(slides_data):
        # Utiliser layout 0 pour la slide 0 et layout 1 pour les suivantes.
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
