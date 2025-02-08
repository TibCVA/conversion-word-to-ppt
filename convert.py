#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys
import os
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt

# --- Débogage : afficher le répertoire courant et la liste des fichiers ---
current_dir = os.path.dirname(os.path.abspath(__file__))
print("Répertoire courant :", current_dir)
print("Fichiers dans ce répertoire :", os.listdir(current_dir))

template_filename = "template_CVA.pptx"
template_path = os.path.join(current_dir, template_filename)
print("Chemin du template :", template_path)

# Vérifier que le fichier existe
if not os.path.exists(template_path):
    print(f"Erreur : Le fichier template n'a pas été trouvé à {template_path}.")
    sys.exit(1)

# Le reste de votre code (exemple simplifié pour la conversion)
def parse_docx_to_slides(doc_path):
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
    p = text_frame.add_paragraph()
    # Détection simple de liste à puces/numérotation par le style (à ajuster selon votre document)
    style_name = paragraph.style.name if paragraph.style else ""
    if "Bullet" in style_name or "Puces" in style_name:
        p.text = "• " + paragraph.text
        p.level = 1
        return p
    elif "Number" in style_name or "Num" in style_name:
        count = counters.get(1, 0) + 1
        counters[1] = count
        p.text = f"{count}. " + paragraph.text
        p.level = 1
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

def clear_all_placeholders(slide):
    for shape in slide.placeholders:
        try:
            shape.text = ""
        except Exception:
            pass

def fill_placeholders(slide, slide_data, slide_index):
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
    create_ppt_from_docx(input_docx, template_path, output_pptx)
