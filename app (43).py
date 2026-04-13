import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

FONT_HEBREW = "David"
FONT_ENGLISH = "Times New Roman"
FONT_SIZE_HEBREW = 13
FONT_SIZE_ENGLISH = 11
FONT_SIZE_SPACING = 11
BULLET_CHAR = "-"

def is_hebrew(text):
    for ch in text:
        if '\u0590' <= ch <= '\u05FF':
            return True
    return False

def is_numbered_paragraph(para):
    """בדוק אם זו רשימה ממוספרת (לא להמיר)"""
    numPr = para._element.find('.//' + qn('w:numPr'))
    if numPr is None:
        return False
    numId = numPr.find(qn('w:numId'))
    ilvl = numPr.find(qn('w:ilvl'))
    if numId is None:
        return False
    try:
        numId_val = numId.get(qn('w:val'))
        ilvl_val = ilvl.get(qn('w:val'), '0') if ilvl is not None else '0'
        numbering = para.part.numbering_part._element
        # מצא את ה-abstractNumId מתוך numId
        for num in numbering.findall(qn('w:num')):
            if num.get(qn('w:numId')) == numId_val:
                abstractNumId = num.find(qn('w:abstractNumId'))
                if abstractNumId is None:
                    break
                absId = abstractNumId.get(qn('w:val'))
                for absNum in numbering.findall(qn('w:abstractNum')):
                    if absNum.get(qn('w:abstractNumId')) == absId:
                        for lvl in absNum.findall(qn('w:lvl')):
                            if lvl.get(qn('w:ilvl')) == ilvl_val:
                                numFmt = lvl.find(qn('w:numFmt'))
                                if numFmt is not None:
                                    fmt = numFmt.get(qn('w:val'), '')
                                    if fmt not in ('bullet',):
                                        return True
    except Exception:
        pass
    return False

def is_bullet_paragraph(para):
    if is_numbered_paragraph(para):
        return False  # מספרים — לא להמיר
    if para._element.find('.//' + qn('w:numPr')) is not None:
        return True
    if para.style and para.style.name and 'List' in para.style.name:
        return True
    return False

def make_paragraph(new_doc, in_table=False):
    p = new_doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if in_table else WD_ALIGN_PARAGRAPH.RIGHT
    pPr = p._element.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        p._element.insert(0, pPr)
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)
    spacing.set(qn('w:line'), '240')
    spacing.set(qn('w:lineRule'), 'auto')
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')
    return p

def set_run_font(run, text, size_pt):
    """קבע פונט כולל cs (עברית) כדי למנוע Cambria"""
    from docx.oxml import OxmlElement
    rPr = run._element.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        run._element.insert(0, rPr)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    font_name = FONT_HEBREW if is_hebrew(text) else FONT_ENGLISH
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)  # חשוב לעברית!
    # גודל
    sz_val = str(size_pt * 2)
    for tag in ['w:sz', 'w:szCs']:
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn('w:val'), sz_val)

def add_run(p, text, bold=False, italic=False, underline=False):
    run = p.add_run(text)
    size = FONT_SIZE_HEBREW if is_hebrew(text) else FONT_SIZE_ENGLISH
    set_run_font(run, text, size)
    if bold: run.bold = True
    if italic: run.italic = True
    if underline: run.underline = True
    return run

def add_spacing(new_doc):
    p = new_doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pPr = p._element.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        p._element.insert(0, pPr)
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:line'), '240')
    spacing.set(qn('w:lineRule'), 'auto')
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')
    pPr.append(spacing)
    run = p.add_run('‏')  # RLM תו כיוון
    set_run_font(run, 'א', FONT_SIZE_SPACING)

def process_document(input_bytes):
    old_doc = Document(io.BytesIO(input_bytes))
    new_doc = Document()

    # העתק numbering definitions מהמסמך הישן לחדש
    try:
        import copy
        old_numbering = old_doc.part.numbering_part._element
        new_numbering = new_doc.part.numbering_part._element
        # מחק את כל ה-abstractNum ו-num הקיימים
        for tag in [qn('w:abstractNum'), qn('w:num')]:
            for el in new_numbering.findall(tag):
                new_numbering.remove(el)
        # העתק מהישן ושנה יישור מספרים לימין
        for tag in [qn('w:abstractNum'), qn('w:num')]:
            for el in old_numbering.findall(tag):
                el_copy = copy.deepcopy(el)
                # שנה lvlJc ל-right
                for lvlJc in el_copy.findall('.//' + qn('w:lvlJc')):
                    lvlJc.set(qn('w:val'), 'right')
                # שנה ind - הזז מספרים לימין
                for ind in el_copy.findall('.//' + qn('w:ind')):
                    ind.set(qn('w:right'), ind.get(qn('w:left'), '720'))
                    ind.attrib.pop(qn('w:left'), None)
                    ind.attrib.pop(qn('w:hanging'), None)
                new_numbering.append(el_copy)
    except Exception:
        pass

    for section in new_doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)

    # הסר פסקה ריקה ראשונה
    for p in new_doc.paragraphs:
        p._element.getparent().remove(p._element)

    last_was_content = False
    last_was_numbered = False
    num_counters = {}  # מונה לכל numId

    for elem in list(old_doc.element.body):
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag

        if tag == 'p':
            from docx.text.paragraph import Paragraph
            para = Paragraph(elem, old_doc)
            full_text = "".join([r.text for r in para.runs]).strip()
            if not full_text:
                continue

            if is_numbered_paragraph(para):
                # העתק הפסקה הממוספרת כמו שהיא — אל תיגע בכלום
                import copy
                if last_was_content and not last_was_numbered:
                    add_spacing(new_doc)
                # הוסף לפני sectPr אם קיים, אחרת append
                body = new_doc.element.body
                sectPr = body.find(qn('w:sectPr'))
                if sectPr is not None:
                    sectPr.addprevious(copy.deepcopy(para._element))
                else:
                    body.append(copy.deepcopy(para._element))
                last_was_content = True
                last_was_numbered = True
                continue

            if last_was_content:
                add_spacing(new_doc)

            if is_bullet_paragraph(para):
                clean = full_text.lstrip('•·-– \t')
                runs = para.runs
                bold = runs[0].bold if runs else False
                italic = runs[0].italic if runs else False
                underline = runs[0].underline if runs else False
                p = make_paragraph(new_doc)
                add_run(p, f"{clean} {BULLET_CHAR}", bold=bold, italic=italic, underline=underline)
            else:
                segments = [{'text': r.text, 'bold': r.bold, 'italic': r.italic, 'underline': r.underline} for r in para.runs if r.text]
                p = make_paragraph(new_doc)
                for s in segments:
                    add_run(p, s['text'], bold=s['bold'], italic=s['italic'], underline=s['underline'])

            last_was_content = True
            last_was_numbered = False

        elif tag == 'tbl':
            from docx.table import Table
            if last_was_content:
                add_spacing(new_doc)
            old_table = Table(elem, old_doc)
            new_table = new_doc.add_table(rows=len(old_table.rows), cols=len(old_table.columns))
            new_table.style = 'Table Grid'
            for r_idx, row in enumerate(old_table.rows):
                for c_idx, cell in enumerate(row.cells):
                    cell_text = cell.text.strip()
                    new_cell = new_table.cell(r_idx, c_idx)
                    cp = new_cell.paragraphs[0]
                    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if cell_text:
                        add_run(cp, cell_text)
            last_was_content = True

    output = io.BytesIO()
    new_doc.save(output)
    output.seek(0)
    return output.read()

# ============================================================
# Streamlit UI
# ============================================================

st.set_page_config(page_title="עיצוב מסמכים", layout="centered")
st.title("🖊️ עיצוב מסמך אוטומטי")
st.markdown("העלה קובץ Word — המסמך יצא מעוצב לפי תקן החברה.")

uploaded = st.file_uploader("בחר קובץ Word", type=["docx"])

if uploaded:
    st.success(f"קובץ נטען: {uploaded.name}")
    if st.button("עצב מסמך"):
        with st.spinner("מעצב..."):
            try:
                result = process_document(uploaded.read())
                st.download_button(
                    label="📥 הורד מסמך מעוצב",
                    data=result,
                    file_name=f"מעוצב_{uploaded.name}",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                st.success("המסמך מוכן!")
            except Exception as e:
                st.error(f"שגיאה: {e}")
                raise e
