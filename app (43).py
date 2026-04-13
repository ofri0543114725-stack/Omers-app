import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy
import io
import re

# ============================================================
# כללי עיצוב קבועים
# ============================================================

FONT_HEBREW = "David"
FONT_ENGLISH = "Times New Roman"
FONT_SIZE_HEBREW = 13
FONT_SIZE_ENGLISH = 11
FONT_SIZE_SPACING = 11  # גודל השורה הריקה בין פסקאות
BULLET_CHAR = "-"
PAGE_WIDTH = Inches(8.27)   # A4
PAGE_HEIGHT = Inches(11.69)  # A4
MARGIN = Inches(1)

# ============================================================

def is_hebrew(text):
    """בדוק אם הטקסט מכיל עברית"""
    for ch in text:
        if '\u0590' <= ch <= '\u05FF':
            return True
    return False

def set_run_font(run, text):
    """קבע פונט לפי שפה"""
    if is_hebrew(text):
        run.font.name = FONT_HEBREW
        run.font.size = Pt(FONT_SIZE_HEBREW)
        # עברית דורשת הגדרה גם ב-XML
        run._element.rPr.rFonts.set(qn('w:ascii'), FONT_HEBREW)
        run._element.rPr.rFonts.set(qn('w:hAnsi'), FONT_HEBREW)
        run._element.rPr.rFonts.set(qn('w:cs'), FONT_HEBREW)
    else:
        run.font.name = FONT_ENGLISH
        run.font.size = Pt(FONT_SIZE_ENGLISH)

def ensure_rpr(run):
    """ודא שיש rPr ל-run"""
    rPr = run._element.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        run._element.insert(0, rPr)
    if not hasattr(run._element, 'rPr') or run._element.rPr is None:
        run._element.rPr = rPr
    return rPr

def is_bullet_paragraph(para):
    """בדוק אם פסקה היא רשימת בולטים"""
    # בדוק numPr (רשימה ממוספרת/בולטים)
    numPr = para._element.find('.//' + qn('w:numPr'))
    if numPr is not None:
        return True
    # בדוק סגנון רשימה
    if para.style and para.style.name and 'List' in para.style.name:
        return True
    return False

def convert_bullet_to_dash(para):
    """המר בולט ל- עם יישור נכון"""
    # הסר numPr
    pPr = para._element.find(qn('w:pPr'))
    if pPr is not None:
        numPr = pPr.find(qn('w:numPr'))
        if numPr is not None:
            pPr.remove(numPr)
        # הסר indent של הרשימה
        ind = pPr.find(qn('w:ind'))
        if ind is not None:
            pPr.remove(ind)

    # קבע סגנון Normal
    para.style = para.part.styles['Normal'] if 'Normal' in [s.name for s in para.part.styles] else para.style

    # קבע יישור ימין
    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # הוסף "- " בתחילת הטקסט
    full_text = "".join([r.text for r in para.runs])
    full_text = full_text.lstrip('•·-– \t')

    # נקה runs קיימים
    for run in para.runs:
        run.text = ""

    # כתוב מחדש עם -
    if para.runs:
        first_run = para.runs[0]
        first_run.text = f"{BULLET_CHAR} {full_text}"
        ensure_rpr(first_run)
        set_run_font(first_run, full_text)
    else:
        new_run = para.add_run(f"{BULLET_CHAR} {full_text}")
        ensure_rpr(new_run)
        set_run_font(new_run, full_text)

def set_paragraph_spacing(para):
    """קבע רווחים: single בתוך, ושורה ריקה בין פסקאות מוגדרת בנפרד"""
    pPr = para._element.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        para._element.insert(0, pPr)

    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)

    # רווח שורות single (240 = 1.0)
    spacing.set(qn('w:line'), '240')
    spacing.set(qn('w:lineRule'), 'auto')
    # אפס רווח לפני ואחרי — הריווח יגיע משורה ריקה
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')

def set_paragraph_rtl(para, in_table=False):
    """קבע כיוון RTL ויישור"""
    pPr = para._element.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        para._element.insert(0, pPr)

    # RTL
    bidi = pPr.find(qn('w:bidi'))
    if bidi is None:
        bidi = OxmlElement('w:bidi')
        pPr.append(bidi)
    bidi.set(qn('w:val'), '1')

    # יישור
    jc = pPr.find(qn('w:jc'))
    if jc is None:
        jc = OxmlElement('w:jc')
        pPr.append(jc)

    if in_table:
        jc.set(qn('w:val'), 'center')
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        jc.set(qn('w:val'), 'right')
        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

def create_spacing_paragraph(doc):
    """צור שורה ריקה David 11 כרווח בין פסקאות"""
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')

    # סגנון
    rPr_para = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), FONT_HEBREW)
    rFonts.set(qn('w:hAnsi'), FONT_HEBREW)
    rFonts.set(qn('w:cs'), FONT_HEBREW)
    rPr_para.append(rFonts)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(FONT_SIZE_SPACING * 2))
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), str(FONT_SIZE_SPACING * 2))
    rPr_para.append(sz)
    rPr_para.append(szCs)
    pPr.append(rPr_para)

    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:line'), '240')
    spacing.set(qn('w:lineRule'), 'auto')
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')
    pPr.append(spacing)
    p.append(pPr)

    return p

def process_runs_in_paragraph(para):
    """עבד את כל ה-runs בפסקה — פונט לפי שפה, שמור bold/italic"""
    for run in para.runs:
        if not run.text.strip():
            continue
        bold = run.bold
        italic = run.italic
        underline = run.underline
        text = run.text
        ensure_rpr(run)
        set_run_font(run, text)
        # שמור bold/italic/underline
        if bold:
            run.bold = True
        if italic:
            run.italic = True
        if underline:
            run.underline = True

def process_document(input_bytes):
    doc = Document(io.BytesIO(input_bytes))

    # קבע שוליים A4
    for section in doc.sections:
        section.page_width = PAGE_WIDTH
        section.page_height = PAGE_HEIGHT
        section.left_margin = MARGIN
        section.right_margin = MARGIN
        section.top_margin = MARGIN
        section.bottom_margin = MARGIN

    # אסוף את כל הפסקאות הראשיות (לא בטבלאות)
    body = doc.element.body
    new_body_elements = []

    # עבור על אלמנטים בגוף המסמך
    for elem in list(body):
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag

        if tag == 'p':
            # פסקה רגילה
            from docx.text.paragraph import Paragraph
            para = Paragraph(elem, doc)

            # דלג על פסקאות ריקות לחלוטין (נוסיף רווחים בעצמנו)
            text = "".join([r.text for r in para.runs]).strip()
            if not text:
                continue

            # המר בולטים
            if is_bullet_paragraph(para):
                convert_bullet_to_dash(para)
            else:
                set_paragraph_rtl(para, in_table=False)

            set_paragraph_spacing(para)
            process_runs_in_paragraph(para)

            new_body_elements.append(('para', elem))

        elif tag == 'tbl':
            # טבלה — עבד פסקאות בתוך תאים
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

    # בנה מחדש את ה-body עם שורות רווח בין פסקאות
    # הסר כל האלמנטים
    for child in list(body):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag in ('p', 'tbl'):
            body.remove(child)

    # הוסף חזרה עם רווחים
    last_was_content = False
    for kind, elem in new_body_elements:
        if kind in ('para', 'table'):
            if last_was_content:
                spacing_p = create_spacing_paragraph(doc)
                body.append(spacing_p)
            body.append(elem)
            last_was_content = True
        else:
            body.append(elem)

    # שמור
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
