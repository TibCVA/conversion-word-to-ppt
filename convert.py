#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script de conversion : Word (input.docx) => PowerPoint (output.pptx)
en utilisant le template template_CVA.pptx.

Usage : python convert.py input.docx output.pptx

Le document Word doit être structuré ainsi :
  - "SLIDE X" pour signaler le début d'une diapositive,
  - "Titre :" pour le titre,
  - "Sous-titre / Message clé :" pour le sous-titre,
  - Le reste (paragraphes, listes...) devient le contenu principal.

Le template PPT (template_CVA.pptx) comporte :
  - Une disposition nommée "Diapositive de titre" pour la slide #0 (couverture),
    avec placeholders "PROJECT TITLE", "CVA Presentation title", "Subtitle".
  - Une disposition nommée "Slide_standard layout" pour les slides suivantes,
    avec placeholders "Click to edit Master title style" (titre),
                       "[Optional subtitle]" (sous-titre),
                       "Modifiez les styles du texte du masque" (contenu).

Ce script remplace ces placeholders par le contenu du Word.
"""

import sys
import os
from docx import Document
from pptx import Presentation
from pptx.util import Pt

# -------------------------------------------------------------------------
# 1) Recherche d'un layout PPT par son "name"
# -------------------------------------------------------------------------
def find_layout_by_name(prs, layout_name):
    """
    Retourne un slide layout de la presentation 'prs' dont le name correspond
    à layout_name, sinon None.
    """
    for layout in prs.slide_layouts:
        if layout.name == layout_name:
            return layout
    return None

# -------------------------------------------------------------------------
# 2) Analyse du Word pour en extraire la liste des slides
# -------------------------------------------------------------------------
def parse_docx_to_slides(doc_path):
    """
    Lit le document Word et renvoie une liste de slides.
    Chaque slide est un dict: {"title": str, "subtitle": str, "blocks": [...] }
    """
    doc = Document(doc_path)
    slides = []
    current_slide = None

    for para in doc.paragraphs:
        raw_text = para.text.strip()
        if not raw_text:
            continue

        # Détection du début d'une diapo : "SLIDE X"
        if raw_text.upper().startswith("SLIDE"):
            # On clôt la slide précédente (si existante)
            if current_slide is not None:
                slides.append(current_slide)
            # On crée la nouvelle
            current_slide = {"title": "", "subtitle": "", "blocks": []}
            continue

        # Détection titre
        if raw_text.startswith("Titre :"):
            if current_slide is not None:
                current_slide["title"] = raw_text[len("Titre :"):].strip()
            continue

        # Détection sous-titre
        if raw_text.startswith("Sous-titre / Message clé :"):
            if current_slide is not None:
                current_slide["subtitle"] = raw_text[len("Sous-titre / Message clé :"):].strip()
            continue

        # Sinon, c'est un paragraphe "contenu"
        if current_slide is not None:
            current_slide["blocks"].append(("paragraph", para))

    # Ne pas oublier le dernier slide s'il y en a un en cours
    if current_slide is not None:
        slides.append(current_slide)

    return slides

# -------------------------------------------------------------------------
# 3) Gestion des listes (puces, numéros) + format runs
# -------------------------------------------------------------------------
def get_list_type(paragraph):
    """
    Détermine si un paragraphe Word est une liste à puces ou numérotée,
    en inspectant son XML interne (bullet ou decimal).
    """
    xml = paragraph._p.xml
    if 'w:numFmt val="bullet"' in xml:
        return "bullet"
    elif 'w:numFmt val="decimal"' in xml:
        return "decimal"
    return None

def get_indentation_level(paragraph):
    """
    Convertit la left_indent (si présente) en un niveau hiérarchique
    approximatif pour la liste (niveau 0, 1, etc.).
    """
    indent = paragraph.paragraph_format.left_indent
    if indent is None:
        return 0
    try:
        # 0.25 pouce = passage de niveau
        return int(indent.inches / 0.25)
    except:
        return 0

def add_paragraph_with_runs(text_frame, paragraph, counters):
    """
    Insère un paragraphe dans un text_frame PPT en respectant :
    - listes à puces / numéros (détectés dans Word),
    - gras/italique/souligné/taille (run-by-run).
    """
    new_p = text_frame.add_paragraph()
    style_liste = get_list_type(paragraph)
    niveau = get_indentation_level(paragraph)

    if style_liste == "bullet":
        # Simple liste à puces
        new_p.text = "• " + paragraph.text
        new_p.level = niveau
        return new_p
    elif style_liste == "decimal":
        # Liste numérotée
        count = counters.get(niveau, 0) + 1
        counters[niveau] = count
        new_p.text = f"{count}. " + paragraph.text
        new_p.level = niveau
        return new_p
    else:
        # Paragraphe normal avec runs
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
                r.font.size = run.font.size  # Taille spécifiée dans Word
            else:
                r.font.size = Pt(14)  # Sinon, taille par défaut

        return new_p

# -------------------------------------------------------------------------
# 4) Fonctions d'insertion dans le PPT
# -------------------------------------------------------------------------
def clear_all_placeholders(slide):
    """
    Supprime (vide) le texte de tous les placeholders de la diapositive,
    afin d'éviter de conserver le texte par défaut du masque.
    """
    for shape in slide.placeholders:
        if shape.has_text_frame:
            shape.text = ""

def fill_placeholders(slide, slide_data, slide_index):
    """
    Remplit les placeholders de la diapo 'slide' selon :
    - slide_index = 0 => diapo de couverture (titre / sous-titre)
    - slide_index >= 1 => diapo standard (titre / sous-titre / contenu).
    """
    if slide_index == 0:
        # Couverture : placeholders "PROJECT TITLE", "CVA Presentation title", "Subtitle"
        for shape in slide.placeholders:
            if shape.has_text_frame:
                placeholder_txt = shape.text.strip().lower()
                if "project title" in placeholder_txt:
                    shape.text = slide_data["title"]
                elif "cva presentation title" in placeholder_txt:
                    shape.text = slide_data["title"]
                elif "subtitle" in placeholder_txt:
                    shape.text = slide_data["subtitle"]
    else:
        # Slides standard
        for shape in slide.placeholders:
            if not shape.has_text_frame:
                continue
            placeholder_txt = shape.text.strip().lower()

            # Titre
            if "click to edit master title style" in placeholder_txt:
                shape.text = slide_data["title"]

            # Sous-titre
            elif "[optional subtitle]" in placeholder_txt:
                shape.text = slide_data["subtitle"]

            # Contenu principal
            elif "modifiez les styles du texte du masque" in placeholder_txt:
                tf = shape.text_frame
                tf.clear()
                list_counters = {}
                for (block_type, block_data) in slide_data["blocks"]:
                    if block_type == "paragraph":
                        add_paragraph_with_runs(tf, block_data, list_counters)
                # Si besoin de gérer tableaux, images, etc. : faire un elif block_type == "table", etc.

# -------------------------------------------------------------------------
# 5) Fonction principale de création de PPT
# -------------------------------------------------------------------------
def create_ppt_from_docx(input_docx, template_path, output_pptx):
    # On récupère la description de chaque slide depuis le Word
    slides_data = parse_docx_to_slides(input_docx)

    # On ouvre le template PPT
    prs = Presentation(template_path)

    # On cherche les layouts attendus par leur "name"
    cover_layout = find_layout_by_name(prs, "Diapositive de titre")
    standard_layout = find_layout_by_name(prs, "Slide_standard layout")

    # Fallback si on ne les trouve pas
    if not cover_layout:
        cover_layout = prs.slide_layouts[0]
    if not standard_layout:
        # On suppose que le 2e layout est le standard
        standard_layout = prs.slide_layouts[1]

    # Pour chaque slide analysé depuis le Word
    for idx, slide_data in enumerate(slides_data):
        # Sélection du layout (0 => couverture, 1+ => standard)
        if idx == 0:
            layout = cover_layout
        else:
            layout = standard_layout

        # On crée la diapo
        slide = prs.slides.add_slide(layout)

        # On vide tous les placeholders, puis on insère le contenu
        clear_all_placeholders(slide)
        fill_placeholders(slide, slide_data, idx)

    # On enregistre le résultat
    prs.save(output_pptx)
    print("Conversion terminée. Fichier généré :", output_pptx)

# -------------------------------------------------------------------------
# 6) Point d'entrée si on exécute ce script directement
# -------------------------------------------------------------------------
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage : python convert.py input.docx output.pptx")
        sys.exit(1)

    input_docx = sys.argv[1]
    output_pptx = sys.argv[2]

    # On détermine où se trouve le template (même dossier que le script)
    script_folder = os.path.dirname(os.path.abspath(__file__))
    template_file = os.path.join(script_folder, "template_CVA.pptx")

    # Exécution
    create_ppt_from_docx(input_docx, template_file, output_pptx)
