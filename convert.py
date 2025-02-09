#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Convertisseur Word vers PowerPoint

Ce script convertit un document Word (.docx) en présentation PowerPoint (.pptx)
en utilisant un template existant comme arrière-plan. Le document Word doit être
structuré avec des marqueurs "SLIDE X" pour délimiter chaque diapositive.

Structure attendue du Word :
    • "SLIDE X" : début d'une nouvelle diapositive
    • "Titre : " : titre de la diapositive
    • "Sous-titre / Message clé : " : sous-titre de la diapositive
    • Contenu : paragraphes suivants (avec puces et numérotation préservées)
    • Tableaux : conservés et positionnés après le texte

Usage:
    python convert.py input.docx output.pptx

Le fichier template_CVA.pptx doit être présent dans le même dossier.
"""

import sys
import os
from docx import Document
from docx.shared import Pt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor

# Conversion pixels vers pouces (96 pixels = 1 pouce)
def px_to_inch(px):
    return float(px) / 96.0

# Définition des zones de positionnement
TITLE_ZONE = {
    "x": Inches(px_to_inch(76)),
    "y": Inches(px_to_inch(35)),
    "width": Inches(max(px_to_inch(1382), 1)),
    "height": Inches(max(px_to_inch(70), 0.5))
}

SUBTITLE_ZONE = {
    "x": Inches(px_to_inch(76)),
    "y": Inches(px_to_inch(119)),
    "width": Inches(max(px_to_inch(1382), 1)),
    "height": Inches(max(px_to_inch(56), 0.5))
}

CONTENT_ZONE = {
    "x": Inches(px_to_inch(76)),
    "y": Inches(px_to_inch(189)),
    "width": Inches(max(px_to_inch(1382), 1)),
    "height": Inches(max(px_to_inch(425), 1))
}

def add_formatted_paragraph(paragraph, content, number_counters=None):
    """
    Ajoute et formate un paragraphe dans une zone de texte PowerPoint.
    
    Args:
        paragraph: Paragraphe PowerPoint à formater
        content: Dictionnaire contenant le texte et les informations de style
        number_counters: Dictionnaire pour suivre la numérotation par niveau
    """
    if number_counters is None:
        number_counters = {}
        
    level = content.get("level", 0)
    prefix = ""
    
    # Gestion des listes (puces et numérotation)
    if content.get("list_type"):
        paragraph.level = level
        indent = "    " * level
        
        if content["list_type"] == "bullet":
            prefix = f"{indent}• "
        else:  # Liste numérotée
            if level not in number_counters:
                number_counters[level] = 0
            number_counters[level] += 1
            # Réinitialisation des compteurs des niveaux supérieurs
            higher_levels = [l for l in number_counters.keys() if l > level]
            for l in higher_levels:
                number_counters[l] = 0
            prefix = f"{indent}{number_counters[level]}. "
    
    # Application du formatage par runs
    runs = content.get("runs", [])
    if runs:
        first_run = True
        for run_format in runs:
            run = paragraph.add_run()
            if first_run:
                run.text = prefix + run_format["text"]
                first_run = False
            else:
                run.text = run_format["text"]
            
            run.font.name = "Arial"
            run.font.size = Pt(11)
            run.font.bold = run_format.get("bold", False)
            run.font.italic = run_format.get("italic", False)
            run.font.underline = run_format.get("underline", False)
    else:
        paragraph.text = prefix + content["text"]
        paragraph.font.name = "Arial"
        paragraph.font.size = Pt(11)

def add_text_box(slide, text, zone, font_size=11, bold=False):
    """
    Ajoute une zone de texte à une slide PowerPoint.
    
    Args:
        slide: Slide PowerPoint
        text: Texte à ajouter
        zone: Dictionnaire avec x, y, width, height en pouces
        font_size: Taille de la police (défaut: 11)
        bold: Texte en gras (défaut: False)
    
    Returns:
        Shape: La forme créée
    """
    shape = slide.shapes.add_textbox(
        zone["x"], zone["y"],
        zone["width"], zone["height"]
    )
    text_frame = shape.text_frame
    text_frame.word_wrap = True
    text_frame.margin_left = text_frame.margin_right = Inches(0.1)
    
    p = text_frame.paragraphs[0]
    p.text = text
    p.font.name = "Arial"
    p.font.size = Pt(font_size)
    p.font.bold = bold
    
    return shape

def add_table(slide, table, left, top, width, height):
    """
    Ajoute un tableau à une slide PowerPoint.
    
    Args:
        slide: Slide PowerPoint
        table: Tableau Word à copier
        left, top: Position en pouces
        width, height: Dimensions en pouces
    
    Returns:
        Shape: Le tableau créé
    """
    shape = slide.shapes.add_table(
        len(table.rows), len(table.columns),
        left, top, width, height
    )
    
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            target_cell = shape.table.cell(i, j)
            target_cell.text = cell.text
            
            for paragraph in target_cell.text_frame.paragraphs:
                paragraph.font.name = "Arial"
                paragraph.font.size = Pt(10)
                if i == 0:  # Première ligne en gras
                    paragraph.font.bold = True
    
    return shape

def create_slide(prs, slide_data):
    """
    Crée une slide PowerPoint complète avec tous ses éléments.
    
    Args:
        prs: Présentation PowerPoint
        slide_data: Dictionnaire contenant les données de la slide
    
    Returns:
        Slide: La slide créée
    """
    # Sélection du layout
    layout = next((l for l in prs.slide_layouts if l.name == 'Blank'), 
                 prs.slide_layouts[0])
    slide = prs.slides.add_slide(layout)
    
    # Ajout du titre
    add_text_box(slide, slide_data["title"], TITLE_ZONE, font_size=22, bold=True)
    
    # Ajout du sous-titre
    add_text_box(slide, slide_data["subtitle"], SUBTITLE_ZONE, font_size=18)
    
    # Zone de contenu
    if slide_data["content"]:
        content_box = slide.shapes.add_textbox(
            CONTENT_ZONE["x"], CONTENT_ZONE["y"],
            CONTENT_ZONE["width"], CONTENT_ZONE["height"]
        )
        text_frame = content_box.text_frame
        text_frame.word_wrap = True
        text_frame.margin_left = text_frame.margin_right = Inches(0.1)
        
        number_counters = {}
        first = True
        
        for content in slide_data["content"]:
            if first:
                p = text_frame.paragraphs[0]
                first = False
            else:
                p = text_frame.add_paragraph()
            add_formatted_paragraph(p, content, number_counters)
        
        # Position Y pour les tableaux
        current_y = CONTENT_ZONE["y"] + content_box.height + Inches(0.2)
    else:
        current_y = CONTENT_ZONE["y"]
    
    # Ajout des tableaux
    for table in slide_data["tables"]:
        if current_y + Inches(1) > (CONTENT_ZONE["y"] + CONTENT_ZONE["height"]):
            print(f"Warning: Espace insuffisant pour un tableau dans la slide")
            break
            
        height = Inches(len(table.rows) * 0.3)
        add_table(
            slide, table,
            CONTENT_ZONE["x"], current_y,
            CONTENT_ZONE["width"], height
        )
        current_y += height + Inches(0.2)
    
    return slide

def parse_word_document(doc_path):
    """
    Parse le document Word et extrait les données de chaque slide.
    
    Args:
        doc_path: Chemin vers le fichier Word
    
    Returns:
        List: Liste des slides avec leur contenu
    """
    doc = Document(doc_path)
    slides_data = []
    current_slide = None
    
    for element in doc.element.body:
        if element.tag.endswith('p'):
            paragraph = doc.paragraphs[len(doc.paragraphs)-1]
            text = paragraph.text.strip()
            
            if not text:
                continue
                
            if text.upper().startswith("SLIDE"):
                if current_slide is not None:
                    slides_data.append(current_slide)
                current_slide = {
                    "title": "",
                    "subtitle": "",
                    "content": [],
                    "tables": []
                }
            elif current_slide is not None:
                if text.startswith("Titre :"):
                    current_slide["title"] = text[len("Titre :"):].strip()
                elif text.startswith("Sous-titre / Message clé :"):
                    current_slide["subtitle"] = text[len("Sous-titre / Message clé :"):].strip()
                else:
                    # Analyse du style du paragraphe
                    style_info = {
                        "text": text,
                        "level": 0,
                        "list_type": None,
                        "runs": []
                    }
                    
                    # Détection du style de liste
                    try:
                        if paragraph._element.pPr is not None:
                            if paragraph._element.pPr.numPr is not None:
                                ilvl = paragraph._element.pPr.numPr.ilvl
                                if ilvl is not None:
                                    style_info["level"] = ilvl.val
                                
                                numId = paragraph._element.pPr.numPr.numId
                                if numId is not None:
                                    if numId.val in [1, 2]:  # Valeurs courantes pour les puces
                                        style_info["list_type"] = "bullet"
                                    else:
                                        style_info["list_type"] = "number"
                    except:
                        pass
                    
                    # Capture du formatage des runs
                    for run in paragraph.runs:
                        run_info = {
                            "text": run.text,
                            "bold": bool(run.bold),
                            "italic": bool(run.italic),
                            "underline": bool(run.underline)
                        }
                        style_info["runs"].append(run_info)
                    
                    current_slide["content"].append(style_info)
                    
        elif element.tag.endswith('tbl'):
            if current_slide is not None:
                # Récupération de l'index correct du tableau
                idx = len([e for e in doc.element.body[:doc.element.body.index(element)]
                         if e.tag.endswith('tbl')])
                current_slide["tables"].append(doc.tables[idx])
    
    # Ajout de la dernière slide
    if current_slide is not None:
        slides_data.append(current_slide)
    
    return slides_data

def main(input_docx, template_pptx, output_pptx):
    """
    Fonction principale de conversion Word vers PowerPoint.
    
    Args:
        input_docx: Chemin du fichier Word source
        template_pptx: Chemin du template PowerPoint
        output_pptx: Chemin du fichier PowerPoint de sortie
    """
    print("\nValidation des fichiers...")
    if not os.path.exists(input_docx):
        raise FileNotFoundError(f"Fichier Word non trouvé: {input_docx}")
    if not os.path.exists(template_pptx):
        raise FileNotFoundError(f"Template PowerPoint non trouvé: {template_pptx}")
    
    print("\nChargement du template PowerPoint...")
    prs = Presentation(template_pptx)
    print("Layouts disponibles:")
    for i, layout in enumerate(prs.slide_layouts):
        print(f" - Layout {i}: {layout.name}")
    
    print("\nExtraction du contenu Word...")
    slides_data = parse_word_document(input_docx)
    
    print("\nCréation de la présentation PowerPoint...")
    # Suppression des slides existantes
    xml_slides = prs.slides._sldIdLst
    for sld in list(xml_slides):
        xml_slides.remove(sld)
    
    # Création des nouvelles slides
    for i, slide_data in enumerate(slides_data, 1):
        print(f"Création de la slide {i}/{len(slides_data)}")
        create_slide(prs, slide_data)
    
    print("\nSauvegarde de la présentation...")
    prs.save(output_pptx)
    print(f"Conversion terminée ! Fichier sauvegardé : {output_pptx}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python convert.py input.docx output.pptx")
        sys.exit(1)
    
    try:
        input_docx = sys.argv[1]
        output_pptx = sys.argv[2]
        template_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template_CVA.pptx")
        
        main(input_docx, template_file, output_pptx)
    except Exception as e:
        print(f"Erreur lors de la