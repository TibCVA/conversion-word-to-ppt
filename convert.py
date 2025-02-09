#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os
from docx import Document
from docx.shared import Pt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor

def px_to_inch(px):
    """Conversion des pixels en pouces"""
    return float(px) / 96.0

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

def is_list_paragraph(paragraph):
    """Détermine si un paragraphe est une liste à puces ou numérotée"""
    try:
        style_name = paragraph.style.name.lower()
        return 'list' in style_name or 'bullet' in style_name or 'number' in style_name
    except:
        return False

def get_paragraph_format(paragraph):
    """Extrait le formatage d'un paragraphe"""
    format_data = {
        "text": paragraph.text,
        "is_list": is_list_paragraph(paragraph),
        "level": 0,
        "runs": []
    }
    
    # Détection du niveau d'indentation
    try:
        if paragraph.paragraph_format.left_indent:
            format_data["level"] = int(paragraph.paragraph_format.left_indent.pt / 36)
    except:
        pass
    
    # Extraction du formatage des runs
    for run in paragraph.runs:
        run_format = {
            "text": run.text,
            "bold": bool(run.bold),
            "italic": bool(run.italic),
            "underline": bool(run.underline)
        }
        format_data["runs"].append(run_format)
    
    return format_data

def create_text_shape(slide, text, left, top, width, height, font_name="Arial", 
                     font_size=11, bold=False, level=0):
    """Crée une forme de texte sur la slide"""
    shape = slide.shapes.add_textbox(left, top, width, height)
    text_frame = shape.text_frame
    text_frame.word_wrap = True
    text_frame.margin_left = text_frame.margin_right = Inches(0.1)
    
    p = text_frame.paragraphs[0]
    
    # Ajout du texte avec indentation si nécessaire
    indent = "    " * level
    p.text = f"{indent}{text}"
    
    # Application des styles
    p.font.name = font_name
    p.font.size = Pt(font_size)
    p.font.bold = bold
    
    return shape

def add_formatted_paragraph(slide, format_data, left, top, width, height):
    """Ajoute un paragraphe formaté à la slide"""
    shape = slide.shapes.add_textbox(left, top, width, height)
    text_frame = shape.text_frame
    text_frame.word_wrap = True
    
    p = text_frame.paragraphs[0]
    p.text = ""
    
    # Ajout des puces si nécessaire
    if format_data["is_list"]:
        indent = "    " * format_data["level"]
        prefix = f"{indent}• "
    else:
        prefix = ""
    
    # Application du formatage par runs
    if format_data["runs"]:
        first = True
        for run_format in format_data["runs"]:
            run = p.add_run()
            if first:
                run.text = prefix + run_format["text"]
                first = False
            else:
                run.text = run_format["text"]
            
            run.font.name = "Arial"
            run.font.size = Pt(11)
            run.font.bold = run_format["bold"]
            run.font.italic = run_format["italic"]
            run.font.underline = run_format["underline"]
    else:
        p.text = prefix + format_data["text"]
        p.font.name = "Arial"
        p.font.size = Pt(11)
    
    return shape

def add_table_to_slide(slide, table, left, top, width, height):
    """Ajoute un tableau à la slide"""
    rows = len(table.rows)
    cols = len(table.columns)
    
    shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    ppt_table = shape.table
    
    # Copie du contenu et formatage
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            target_cell = ppt_table.cell(i, j)
            target_cell.text = cell.text
            
            # Application du style
            for paragraph in target_cell.text_frame.paragraphs:
                paragraph.font.name = "Arial"
                paragraph.font.size = Pt(10)
                if i == 0:  # En-tête en gras
                    paragraph.font.bold = True
    
    return shape

def parse_word_document(doc_path):
    """Parse le document Word et extrait les slides avec leur contenu"""
    doc = Document(doc_path)
    slides_data = []
    current_slide = None
    all_elements = []
    current_slide_index = -1
    
    # Collecte tous les éléments (paragraphes et tableaux) dans l'ordre
    for element in doc.paragraphs:
        all_elements.append(('paragraph', element))
    for table in doc.tables:
        # On insère le tableau à sa position relative par rapport aux paragraphes
        table_index = doc.element.body.index(table._element)
        # Trouver la bonne position dans all_elements
        insert_pos = 0
        for i, (elem_type, elem) in enumerate(all_elements):
            if elem_type == 'paragraph':
                if doc.element.body.index(elem._element) > table_index:
                    insert_pos = i
                    break
        all_elements.insert(insert_pos, ('table', table))
    
    # Traitement des éléments dans l'ordre
    for elem_type, element in all_elements:
        if elem_type == 'paragraph':
            text = element.text.strip()
            if not text:
                continue
            
            if text.upper().startswith("SLIDE"):
                current_slide_index += 1
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
                    # Analyse détaillée du formatage
                    para_format = {
                        "text": text,
                        "style": element.style.name,
                        "level": 0,
                        "list_type": None,
                        "runs": []
                    }
                    
                    # Détection du type de liste et du niveau
                    if element._element.pPr is not None:
                        if element._element.pPr.numPr is not None:
                            ilvl = element._element.pPr.numPr.ilvl
                            if ilvl is not None:
                                para_format["level"] = ilvl.val
                            
                            # Détermine si c'est une puce ou une numérotation
                            numId = element._element.pPr.numPr.numId
                            if numId is not None:
                                if numId.val in [1, 2]:  # Valeurs courantes pour les puces
                                    para_format["list_type"] = "bullet"
                                else:
                                    para_format["list_type"] = "number"
                    
                    # Capture du formatage des runs
                    for run in element.runs:
                        run_format = {
                            "text": run.text,
                            "bold": bool(run.bold),
                            "italic": bool(run.italic),
                            "underline": bool(run.underline)
                        }
                        para_format["runs"].append(run_format)
                    
                    current_slide["content"].append(para_format)
                    
        elif elem_type == 'table' and current_slide is not None:
            current_slide["tables"].append(element)
    
    # Ajout de la dernière slide
    if current_slide is not None:
        slides_data.append(current_slide)
    
    return slides_data

def create_slide(prs, slide_data):
    """Crée une slide avec son contenu"""
    # Sélection du layout
    layout = next((layout for layout in prs.slide_layouts if layout.name == 'Blank'), 
                 prs.slide_layouts[0])
    
    slide = prs.slides.add_slide(layout)
    
    # Ajout du titre
    create_text_shape(
        slide, slide_data["title"],
        TITLE_ZONE["x"], TITLE_ZONE["y"],
        TITLE_ZONE["width"], TITLE_ZONE["height"],
        font_size=22, bold=True
    )
    
    # Ajout du sous-titre
    create_text_shape(
        slide, slide_data["subtitle"],
        SUBTITLE_ZONE["x"], SUBTITLE_ZONE["y"],
        SUBTITLE_ZONE["width"], SUBTITLE_ZONE["height"],
        font_size=18
    )
    
    # Création d'une unique zone de contenu pour tous les paragraphes
    if slide_data["content"]:
        content_shape = slide.shapes.add_textbox(
            CONTENT_ZONE["x"], CONTENT_ZONE["y"],
            CONTENT_ZONE["width"], CONTENT_ZONE["height"]
        )
        text_frame = content_shape.text_frame
        text_frame.word_wrap = True
        text_frame.margin_left = text_frame.margin_right = Inches(0.1)
        
        # Dictionnaire pour suivre les compteurs de numérotation par niveau
        number_counters = {}
        
        # Ajout de chaque paragraphe dans la même zone de texte
        first_paragraph = True
        for content in slide_data["content"]:
            if not first_paragraph:
                paragraph = text_frame.add_paragraph()
            else:
                paragraph = text_frame.paragraphs[0]
                first_paragraph = False
            
            format_paragraph_content(paragraph, text_frame, content, number_counters)
    
    # Position pour les tableaux
    current_y = CONTENT_ZONE["y"]
    if slide_data["content"]:
        current_y += content_shape.height + Inches(0.2)
    
    # Ajout des tableaux
    for table in slide_data["tables"]:
        if current_y + Inches(1) > (CONTENT_ZONE["y"] + CONTENT_ZONE["height"]):
            print(f"Warning: Espace insuffisant pour un tableau dans la slide")
            break
        
        height = Inches(len(table.rows) * 0.3)
        shape = add_table_to_slide(
            slide, table,
            CONTENT_ZONE["x"], current_y,
            CONTENT_ZONE["width"], height
        )
        current_y += height + Inches(0.2)
    
    return slide
    
    # Position pour les tableaux
    current_y = CONTENT_ZONE["y"]
    if slide_data["content"]:
        # Ajuster current_y en fonction de la hauteur du contenu texte
        current_y += content_shape.height + Inches(0.2)
    
    # Ajout des tableaux
    for table in slide_data["tables"]:
        if current_y + Inches(1) > (CONTENT_ZONE["y"] + CONTENT_ZONE["height"]):
            print(f"Warning: Espace insuffisant pour un tableau dans la slide")
            break
        
        height = Inches(len(table.rows) * 0.3)
        shape = add_table_to_slide(
            slide, table,
            CONTENT_ZONE["x"], current_y,
            CONTENT_ZONE["width"], height
        )
        current_y += height + Inches(0.2)
    
    return slide

def create_ppt_from_docx(input_docx, template_pptx, output_pptx):
    """Fonction principale de conversion"""
    print("\nValidation des fichiers...")
    if not os.path.exists(input_docx):
        raise FileNotFoundError(f"Le fichier Word {input_docx} n'existe pas.")
    if not os.path.exists(template_pptx):
        raise FileNotFoundError(f"Le template PowerPoint {template_pptx} n'existe pas.")
    
    print("\nChargement du template PowerPoint...")
    prs = Presentation(template_pptx)
    print("Layouts disponibles dans le template:")
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
    print(f"Conversion terminée ! Le fichier a été sauvegardé sous : {output_pptx}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python convert.py input.docx output.pptx")
        sys.exit(1)
    
    try:
        input_docx = sys.argv[1]
        output_pptx = sys.argv[2]
        template_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template_CVA.pptx")
        
        create_ppt_from_docx(input_docx, template_file, output_pptx)
    except Exception as e:
        print(f"Erreur lors de la conversion : {str(e)}")
        sys.exit(1)