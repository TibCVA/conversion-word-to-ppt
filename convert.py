import sys
import os
from docx import Document
from pptx import Presentation
from pptx.util import Pt
import logging

# Configuration du logging
logging.basicConfig(level=logging.INFO)

def parse_docx_to_slides(doc_path):
    doc = Document(doc_path)
    slides_data = []
    current_slide = None

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        if text.upper().startswith("SLIDE"):
            if current_slide is not None:
                slides_data.append(current_slide)
            current_slide = {"title": "", "subtitle": "", "blocks": []}
            continue
        if text.startswith("Titre :"):
            if current_slide is not None:
                current_slide["title"] = text[len("Titre :"):].strip()
            continue
        if text.startswith("Sous-titre / Message clé :"):
            if current_slide is not None:
                current_slide["subtitle"] = text[len("Sous-titre / Message clé :"):].strip()
            continue
        if current_slide is not None:
            current_slide["blocks"].append(("paragraph", para))
    if current_slide is not None:
        slides_data.append(current_slide)

    logging.info("Contenu extrait du Word:")
    for i, slide in enumerate(slides_data):
        logging.info(f"Slide {i} - Titre: {slide['title']} | Sous-titre: {slide['subtitle']} | Nb de blocs: {len(slide['blocks'])}")
    return slides_data

def get_list_type(paragraph):
    xml = paragraph._p.xml
    if 'w:numFmt val="bullet"' in xml:
        return "bullet"
    elif 'w:numFmt val="decimal"' in xml:
        return "decimal"
    return None

def get_indentation_level(paragraph):
    indent = paragraph.paragraph_format.left_indent
    if indent is None:
        return 0
    try:
        return int(indent.inches / 0.25)
    except:
        return 0

def add_paragraph_with_runs(text_frame, paragraph, counters):
    new_p = text_frame.add_paragraph()
    style_list = get_list_type(paragraph)
    level = get_indentation_level(paragraph)
    if style_list == "bullet":
        new_p.text = "• " + paragraph.text
        new_p.level = level
    elif style_list == "decimal":
        count = counters.get(level, 0) + 1
        counters[level] = count
        new_p.text = f"{count}. " + paragraph.text
        new_p.level = level
    else:
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
                r.font.size = run.font.size
            else:
                r.font.size = Pt(14)
    return new_p

def fill_cover_slide(slide, slide_data):
    logging.info("Placeholders dans la slide de couverture:")
    for ph in slide.placeholders:
        logging.info(f" - Nom: {ph.name}, Texte: '{ph.text}'")

    # Utilisation des noms des placeholders
    for ph in slide.placeholders:
        if "PROJECT TITLE" in ph.text:
            ph.text = slide_data["title"]
            logging.info("Placeholder 'PROJECT TITLE' mis à jour avec le titre:", slide_data["title"])
        elif "CVA Presentation title" in ph.text:
            ph.text = slide_data["subtitle"]
            logging.info("Placeholder 'CVA Presentation title' mis à jour avec le sous-titre:", slide_data["subtitle"])

def fill_standard_slide(slide, slide_data):
    logging.info("Placeholders dans une slide standard:")
    for ph in slide.placeholders:
        logging.info(f" - Nom: {ph.name}, Texte: '{ph.text}'")

    # Utilisation des noms des placeholders
    for ph in slide.placeholders:
        if "Click to edit Master title style" in ph.text:
            ph.text = slide_data["title"]
            logging.info("Placeholder 'Click to edit Master title style' mis à jour avec le titre:", slide_data["title"])
        elif "[Optional subtitle]" in ph.text:
            ph.text = slide_data["subtitle"]
            logging.info("Placeholder '[Optional subtitle]' mis à jour avec le sous-titre:", slide_data["subtitle"])
        elif "Modifiez les styles du texte du masque" in ph.text:
            ph.text_frame.clear()
            counters = {}
            for (block_type, block_para) in slide_data["blocks"]:
                if block_type == "paragraph":
                    add_paragraph_with_runs(ph.text_frame, block_para, counters)
            logging.info("Placeholder 'Modifiez les styles du texte du masque' rempli avec le contenu.")

def create_ppt_from_docx(input_docx, template_pptx, output_pptx):
    slides_data = parse_docx_to_slides(input_docx)
    prs = Presentation(template_pptx)

    cover_layout = None
    standard_layout = None
    for layout in prs.slide_layouts:
        if layout.name == "Diapositive de titre":
            cover_layout = layout
        elif layout.name == "Slide_standard layout":
            standard_layout = layout
    if not cover_layout:
        cover_layout = prs.slide_layouts[0]
    if not standard_layout:
        standard_layout = prs.slide_layouts[1]

    if slides_data:
        slide0 = prs.slides.add_slide(cover_layout)
        logging.info("=== Traitement de la slide 0 ===")
        fill_cover_slide(slide0, slides_data[0])
    else:
        logging.warning("Aucune slide trouvée dans le document Word.")

    for idx, slide_data in enumerate(slides_data[1:], start=1):
        slide = prs.slides.add_slide(standard_layout)
        logging.info(f"=== Traitement de la slide {idx} ===")
        fill_standard_slide(slide, slide_data)

    prs.save(output_pptx)
    logging.info("Conversion terminée :", output_pptx)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage : python convert.py input.docx output.pptx")
        sys.exit(1)
    input_docx = sys.argv[1]
    output_pptx = sys.argv[2]
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_file = os.path.join(script_dir, "template_CVA.pptx")
    create_ppt_from_docx(input_docx, template_file, output_pptx)
