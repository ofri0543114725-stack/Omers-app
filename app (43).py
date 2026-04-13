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
PAGE_WIDTH = Inches(8.27)
PAGE_HEIGHT = Inches(11.69)
MARGIN = Inches(1)

def is_hebrew(text):
    for ch in text:
        if '\u0590' <= ch <= '\u05FF':
            return True
    return False

def ensure_rpr(run):
    rPr = run._element.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        run._element.insert(0, rPr)
    return rPr

def set_run_font(run, text):
    rPr = ensure_rpr(run)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    if is_hebrew(text):
        rFonts.set(qn('w:ascii'), FONT_HEBREW)
        rFonts.set(qn('w:hAnsi'), FONT_HEBREW)
        rFonts.set(qn('w:cs'), FONT_HEBREW)
        sz_val = FONT_SIZE_HEBREW * 2
    else:
        rFonts.set(qn('w:ascii'), FONT_ENGLISH)
        rFonts.set(qn('w:hAnsi'), FONT_ENGLISH)
        rFonts.set(qn('w:cs'), FONT_ENGLISH)
        sz_val = FONT_SIZE_ENGLISH * 2
    for tag_name in ['w:sz', 'w:szCs']:
        el = rPr.find(qn(tag_name))
        if el is None:
            el = OxmlElement(tag_name)
            rPr.append(el)
        el.set(qn('w:val'), str(sz_val))

def get_or_create_pPr(para):
    pPr = para._element.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        para._element.insert(0, pPr)
    return pPr

def set_paragraph_rtl(para, in_table=False):
    pPr = get_or_create_pPr(para)
    # הסר bidi ו-jc קיימים
    for tag in [qn('w:jc'), qn('w:bidi')]:
        el = pPr.find(tag)
        if el is not None:
            pPr.remove(el)
    # הוסף bidi
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    # הוסף jc
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'center' if in_table else 'right')
    pPr.append(jc)

def set_paragraph_spacing(para):
    pPr = get_or_create_pPr(para)
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)
    spacing.set(qn('w:line'), '240')
    spacing.set(qn('w:lineRule'), 'auto')
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')

def set_all_styles_rtl(doc):
    """קבע jc=right על כל הסגנונות במסמך כדי שלא ידרסו את הפסקאות"""
    styles_elem = doc.element.find(qn('w:styles'))
    if styles_elem is None:
        return
    for style in styles_elem.findall(qn('w:style')):
        pPr = style.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            style.append(pPr)
        # עדכן bidi ו-jc בכל סגנון
        for tag in [qn('w:jc'), qn('w:bidi')]:
            el = pPr.find(tag)
            if el is not None:
                pPr.remove(el)
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '1')
        pPr.append(bidi)
        jc = OxmlElement('w:jc')
        jc.set(qn('w:val'), 'right')
        pPr.append(jc)
        # עדכן גם rPr של הסגנון לפונט David 13
        rPr = style.find(qn('w:rPr'))
        if rPr is None:
            rPr = OxmlElement('w:rPr')
            style.append(rPr)
        for tag in [qn('w:sz'), qn('w:szCs'), qn('w:rFonts')]:
            el = rPr.find(tag)
            if el is not None:
                rPr.remove(el)
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), FONT_HEBREW)
        rFonts.set(qn('w:hAnsi'), FONT_HEBREW)
        rFonts.set(qn('w:cs'), FONT_HEBREW)
        rPr.insert(0, rFonts)
        sz = OxmlElement('w:sz')
        sz.set(qn('w:val'), str(FONT_SIZE_HEBREW * 2))
        rPr.append(sz)
        szCs = OxmlElement('w:szCs')
        szCs.set(qn('w:val'), str(FONT_SIZE_HEBREW * 2))
        rPr.append(szCs)

def is_bullet_paragraph(para):
    if para._element.find('.//' + qn('w:numPr')) is not None:
        return True
    if para.style and para.style.name and 'List' in para.style.name:
        return True
    return False

def convert_bullet_to_dash(para):
    pPr = get_or_create_pPr(para)
    for tag in [qn('w:numPr'), qn('w:ind')]:
        el = pPr.find(tag)
        if el is not None:
            pPr.remove(el)
    runs_props = [{'bold': r.bold, 'italic': r.italic, 'underline': r.underline} for r in para.runs]
    full_text = "".join([r.text for r in para.runs]).lstrip('•·-– \t')
    for run in para.runs:
        run.text = ""
    if para.runs:
        first_run = para.runs[0]
        first_run.text = f"{BULLET_CHAR} {full_text}"
        props = runs_props[0] if runs_props else {}
        if props.get('bold'): first_run.bold = True
        if props.get('italic'): first_run.italic = True
        if props.get('underline'): first_run.underline = True
        set_run_font(first_run, full_text)
    else:
        new_run = para.add_run(f"{BULLET_CHAR} {full_text}")
        set_run_font(new_run, full_text)
    set_paragraph_rtl(para, in_table=False)

def process_runs_in_paragraph(para):
    for run in para.runs:
        if not run.text.strip():
            continue
        bold, italic, underline = run.bold, run.italic, run.underline
        set_run_font(run, run.text)
        if bold: run.bold = True
        if italic: run.italic = True
        if underline: run.underline = True

def create_spacing_paragraph(doc):
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    rPr_para = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), FONT_HEBREW)
    rFonts.set(qn('w:hAnsi'), FONT_HEBREW)
    rFonts.set(qn('w:cs'), FONT_HEBREW)
    rPr_para.append(rFonts)
    for tag_name, val in [('w:sz', FONT_SIZE_SPACING * 2), ('w:szCs', FONT_SIZE_SPACING * 2)]:
        el = OxmlElement(tag_name)
        el.set(qn('w:val'), str(val))
        rPr_para.append(el)
    pPr.append(rPr_para)
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:line'), '240')
    spacing.set(qn('w:lineRule'), 'auto')
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')
    pPr.append(spacing)
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'right')
    pPr.append(jc)
    p.append(pPr)
    return p

def process_document(input_bytes):
    doc = Document(io.BytesIO(input_bytes))

    # קבע RTL על כל הסגנונות — זה מונע דריסה
    set_all_styles_rtl(doc)

    for section in doc.sections:
        section.page_width = PAGE_WIDTH
        section.page_height = PAGE_HEIGHT
        section.left_margin = MARGIN
        section.right_margin = MARGIN
        section.top_margin = MARGIN
        section.bottom_margin = MARGIN

    body = doc.element.body
    new_body_elements = []

    for elem in list(body):
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag

        if tag == 'p':
            from docx.text.paragraph import Paragraph
            para = Paragraph(elem, doc)
            text = "".join([r.text for r in para.runs]).strip()
            if not text:
                continue
            if is_bullet_paragraph(para):
                convert_bullet_to_dash(para)
            else:
                set_paragraph_rtl(para, in_table=False)
            set_paragraph_spacing(para)
            process_runs_in_paragraph(para)
            new_body_elements.append(('para', elem))

        elif tag == 'tbl':
            from docx.table import Table
            table = Table(elem, doc)
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        set_paragraph_rtl(para, in_table=True)
                        set_paragraph_spacing(para)
                        process_runs_in_paragraph(para)
            new_body_elements.append(('table', elem))

        else:
            new_body_elements.append(('other', elem))

    for child in list(body):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag in ('p', 'tbl'):
            body.remove(child)

    last_was_content = False
    for kind, elem in new_body_elements:
        if kind in ('para', 'table'):
            if last_was_content:
                body.append(create_spacing_paragraph(doc))
            body.append(elem)
            last_was_content = True
        else:
            body.append(elem)

    output = io.BytesIO()
    doc.save(output)
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
