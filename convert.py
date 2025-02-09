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

# Conversion cm en pouces
def cm_to_inch(cm):
    return cm / 2.54

# Définition des zones de positionnement
TITLE_ZONE = {
    "x": Inches(px_to_inch(76)),
    "y": Inches(px_to_inch(35) + 0.4),  # Déplacé de 0.4 pouce vers le bas
    "width": Inches(cm_to_inch(30.33)),  # 30.33 cm
    "height": Inches(max(px_to_inch(70), 0.5))
}

SUBTITLE_ZONE = {
    "x": Inches(px_to_inch(76)),
    "y": Inches(px_to_inch(119)),
    "width": Inches(cm_to_inch(30.33)),  # 30.33 cm
    "height": Inches(max(px_to_inch(56), 0.5))
}

CONTENT_ZONE = {
    "x": Inches(px_to_inch(76)),
    "y": Inches(px_to_inch(189)),
    "width": Inches(cm_to_inch(30.33)),  # 30.33 cm
    "height": Inches(max(px_to_inch(425), 1))
}

def parse_word_document(doc_path):
    """Parse le document Word et extrait les données de chaque slide."""
    doc = Document(doc_path)
    slides_data = []
    current_slide = None
    table_count = 0
    
    # Collecte des paragraphes et tableaux en préservant l'ordre
    elements = []
    for para in doc.paragraphs:
        elements.append(('paragraph', para))
        text = para.text.strip()
        
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
            table_count = 0
        elif current_slide is not None:
            if text.startswith("Titre :"):
                current_slide["title"] = text[len("Titre :"):].strip()
            elif text.startswith("Sous-titre / Message clé :"):
                current_slide["subtitle"] = text[len("Sous-titre / Message clé :"):].strip()
            else:
                style_info = {
                    "text": text,
                    "level": 0,
                    "list_type": None,
                    "runs": []
                }
                
                # Détection des listes
                if para._element.pPr is not None:
                    if para._element.pPr.numPr is not None:
                        numPr = para._element.pPr.numPr
                        ilvl = numPr.ilvl
                        if ilvl is not None:
                            style_info["level"] = ilvl.val
                        numId = numPr.numId
                        if numId is not None:
                            if numId.val in [1, 2]:  # Puces
                                style_info["list_type"] = "bullet"
                            else:  # Numérotation
                                style_info["list_type"] = "number"
                
                # Capture du formatage
                for run in para.runs:
                    run_info = {
                        "text": run.text,
                        "bold": bool(run.bold),
                        "italic": bool(run.italic),
                        "underline": bool(run.underline)
                    }
                    style_info["runs"].append(run_info)
                
                current_slide["content"].append(style_info)
    
    # Traitement des tableaux
    for table in doc.tables:
        if current_slide is not None:
            current_slide["tables"].append(table)
    
    # Ajout de la dernière slide
    if current_slide is not None:
        slides_data.append(current_slide)
    
    return slides_data

def add_formatted_text(paragraph, content, number_counters=None):
    """Ajoute du texte formaté à un paragraphe."""
    if number_counters is None:
        number_counters = {}
    
    level = content.get("level", 0)
    prefix = ""
    
    # Configuration de base du paragraphe
    paragraph.alignment = PP_ALIGN.CENTER  # Centrage du texte
    paragraph.space_after = Pt(3)  # Espacement de 3 points après le paragraphe
    
    # Application des listes
    if content.get("list_type"):
        paragraph.level = level
        indent = "    " * level
        
        if content["list_type"] == "bullet":
            prefix = f"{indent}• "
        else:
            if level not in number_counters:
                number_counters[level] = 0
            number_counters[level] += 1
            
            # Réinitialisation des niveaux supérieurs
            for l in list(number_counters.keys()):
                if l > level:
                    number_counters[l] = 0
                    
            prefix = f"{indent}{number_counters[level]}. "
    
    # Ajout du texte avec formatage
    if content.get("runs"):
        first_run = True
        for run_format in content["runs"]:
            run = paragraph.add_run()
            if first_run:
                run.text = prefix + run_format["text"]
                first_run = False
            else:
                run.text = run_format["text"]
            
            run.font.name = "Arial"
            run.font.size = Pt(12)  # Taille de police augmentée à 12
            run.font.bold = run_format.get("bold", False)
            run.font.italic = run_format.get("italic", False)
            run.font.underline = run_format.get("underline", False)
    else:
        paragraph.text = prefix + content["text"]
        paragraph.font.name = "Arial"
        paragraph.font.size = Pt(12)  # Taille de police augmentée à 12

def create_slide(prs, slide_data):
    """Crée une slide complète avec son contenu."""
    # Sélection du layout
    layout = next((l for l in prs.slide_layouts if l.name == 'Blank'), 
                 prs.slide_layouts[0])
    slide = prs.slides.add_slide(layout)
    
    # Ajout du titre
    title_box = slide.shapes.add_textbox(
        TITLE_ZONE["x"], TITLE_ZONE["y"],
        TITLE_ZONE["width"], TITLE_ZONE["height"]
    )
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    title_frame.paragraphs[0].text = slide_data["title"]
    title_frame.paragraphs[0].font.name = "Arial"
    title_frame.paragraphs[0].font.size = Pt(22)
    title_frame.paragraphs[0].font.bold = True
    
    # Ajout du sous-titre
    subtitle_box = slide.shapes.add_textbox(
        SUBTITLE_ZONE["x"], SUBTITLE_ZONE["y"],
        SUBTITLE_ZONE["width"], SUBTITLE_ZONE["height"]
    )
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.word_wrap = True
    subtitle_frame.paragraphs[0].text = slide_data["subtitle"]
    subtitle_frame.paragraphs[0].font.name = "Arial"
    subtitle_frame.paragraphs[0].font.size = Pt(18)
    
    # Zone de contenu unique
    if slide_data["content"]:
        content_box = slide.shapes.add_textbox(
            CONTENT_ZONE["x"], CONTENT_ZONE["y"],
            CONTENT_ZONE["width"], CONTENT_ZONE["height"]
        )
        content_frame = content_box.text_frame
        content_frame.word_wrap = True
        
        number_counters = {}
        first_paragraph = True
        
        for content in slide_data["content"]:
            if first_paragraph:
                paragraph = content_frame.paragraphs[0]
                first_paragraph = False
            else:
                paragraph = content_frame.add_paragraph()
            
            add_formatted_text(paragraph, content, number_counters)
        
        # Position pour les tableaux
        current_y = CONTENT_ZONE["y"] + content_box.height + Inches(0.2)
    else:
        current_y = CONTENT_ZONE["y"]
    
    # Ajout des tableaux
    for table in slide_data["tables"]:
        if current_y + Inches(1) > (CONTENT_ZONE["y"] + CONTENT_ZONE["height"]):
            print(f"Warning: Espace insuffisant pour un tableau")
            break
        
        slide_table = slide.shapes.add_table(
            len(table.rows), len(table.columns),
            CONTENT_ZONE["x"], current_y,
            CONTENT_ZONE["width"], Inches(len(table.rows) * 0.3)
        ).table
        
        # Copie du contenu
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                target_cell = slide_table.cell(i, j)
                target_cell.text = cell.text
                
                # Formatage
                for paragraph in target_cell.text_frame.paragraphs:
                    paragraph.font.name = "Arial"
                    paragraph.font.size = Pt(10)
                    if i == 0:  # En-tête
                        paragraph.font.bold = True
        
        current_y += Inches(len(table.rows) * 0.3) + Inches(0.2)
    
    return slide

def main(input_docx, template_pptx, output_pptx):
    """Fonction principale de conversion."""
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
        print(f"Erreur lors de la conversion : {str(e)}")
        sys.exit(1)
