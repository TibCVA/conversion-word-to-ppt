from docx import Document
from docx.oxml.ns import qn
from pptx import Presentation
from pptx.util import Inches

# Chemins des fichiers
docx_path = "input.docx"
pptx_template_path = "template.pptx"
pptx_output_path = "output.pptx"

# Ouvrir le document Word et la présentation PPT template
doc = Document(docx_path)
prs = Presentation(pptx_template_path)

# Choisir un layout de diapositive approprié (par ex, index 1 pour "Titre et Contenu")
content_layout = prs.slide_layouts[1]

slide = None  # la diapositive en cours
content_ph = None  # placeholder de contenu en cours

for para in doc.paragraphs:
    style_name = para.style.name or ""
    text = para.text.strip()
    # Si paragraphe de titre (ex: Heading 1) -> nouvelle diapositive
    if style_name.startswith("Heading") or style_name.startswith("Titre"):  
        # Créer une nouvelle slide avec titre
        slide = prs.slides.add_slide(content_layout)
        # Insérer le titre
        title_ph = slide.shapes.title
        title_ph.text = text
        # Récupérer le placeholder de contenu et le vider
        content_ph = slide.placeholders[1]  # placeholder de contenu (idx 1 généralement)
        text_frame = content_ph.text_frame
        text_frame.clear()
    elif style_name.startswith("List") or style_name.startswith("Puces") or style_name.startswith("Num"):
        # Cas d'une liste à puces/num - s'assure qu'une slide existe
        if slide is None:
            slide = prs.slides.add_slide(content_layout)
            content_ph = slide.placeholders[1]
            text_frame = content_ph.text_frame
            text_frame.clear()
        # Ajouter un paragraphe pour cet item de liste
        p = text_frame.add_paragraph()
        p.text = text
        # Déterminer puce ou numérotation via le nom de style
        if "Bullet" in style_name or "Puces" in style_name:
            p.level = 1  # niveau 1 avec puce (défini dans le masque PPT)
        elif "Number" in style_name or "Num" in style_name:
            p.level = 2  # niveau 2 avec numérotation
        else:
            # Style "List Paragraph" ou autre indéfini -> par défaut puce niveau 1
            p.level = 1
    elif text == "":
        # Paragraphe vide (saut de ligne) -> on peut ajouter une ligne vide dans PPT
        if slide and content_ph:
            text_frame.add_paragraph()  # ajoute un paragraphe vide (espacement)
        continue
    else:
        # Paragraphe normal (non liste)
        if slide is None:
            slide = prs.slides.add_slide(content_layout)
            content_ph = slide.placeholders[1]
            text_frame = content_ph.text_frame
            text_frame.clear()
        p = text_frame.add_paragraph()
        p.text = text
        p.level = 0  # texte normal (niveau 0 = pas de puce, selon le masque)

# Parcourir les tableaux du document Word
for table in doc.tables:
    # Créer une nouvelle diapositive pour le tableau (ou adapter si besoin sur la même slide)
    slide = prs.slides.add_slide(content_layout)
    title_ph = slide.shapes.title
    title_ph.text = "Tableau"  # on peut éventuellement utiliser un texte du doc comme titre
    content_ph = slide.placeholders[1]
    # Insérer le tableau aux coordonnées du placeholder de contenu
    left, top, width, height = content_ph.left, content_ph.top, content_ph.width, content_ph.height
    rows, cols = len(table.rows), len(table.columns)
    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    ppt_table = table_shape.table

    # Fixer les largeurs de colonne d'après Word (si disponibles)
    for j, column in enumerate(table.columns):
        if column.width:
            ppt_table.columns[j].width = column.width
    # Recopie du contenu des cellules et gestion des fusions
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            ppt_cell = ppt_table.cell(i, j)
            ppt_cell.text = cell.text.strip()
            # Détecter fusion horizontale (gridSpan) et verticale (vMerge) sur la cellule Word
            tc = cell._tc  # élément XML de la cellule
            grid_span_elems = tc.xpath('.//w:gridSpan')
            v_merge_elems = tc.xpath('.//w:vMerge')
            # Valeurs par défaut
            h_span = 1
            v_span = 1
            if grid_span_elems:
                # nombre de colonnes fusionnées
                span_val = int(grid_span_elems[0].get(qn('w:val')))
                if span_val > 1:
                    h_span = span_val
            if v_merge_elems:
                v_merge_val = v_merge_elems[0].get(qn('w:val'))
                if v_merge_val == "restart":
                    # calculer combien de lignes sont fusionnées verticalement depuis cette cellule
                    for k in range(i+1, rows):
                        nxt_tc = table.cell(k, j)._tc
                        nxt_vmerge = nxt_tc.xpath('.//w:vMerge')
                        if not nxt_vmerge:
                            break
                        nxt_vmerge_val = nxt_vmerge[0].get(qn('w:val'))
                        if nxt_vmerge_val is None or nxt_vmerge_val == "continue":
                            v_span += 1
                        else:
                            break
                else:
                    # si vMerge = continue, on n’agit pas ici (fusion gérée par la cellule 'restart')
                    continue
            # Si une fusion est détectée, fusionner les cellules correspondantes dans le tableau PPT
            if v_span > 1 or h_span > 1:
                end_row = i + v_span - 1
                end_col = j + h_span - 1
                try:
                    ppt_table.cell(i, j).merge(ppt_table.cell(end_row, end_col))
                except IndexError:
                    # sécurité en cas d'indice hors limite (cas improbable si calculs corrects)
                    pass

# Enregistrer la présentation générée
prs.save(pptx_output_path)
