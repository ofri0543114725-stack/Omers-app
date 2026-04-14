import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io, copy, zipfile, re

FONT_HEBREW = "David"
FONT_ENGLISH = "Times New Roman"
FONT_SIZE_HEBREW = 13
FONT_SIZE_ENGLISH = 11
PAGE_WIDTH = Inches(8.27)
PAGE_HEIGHT = Inches(11.69)
MARGIN = Inches(1)

def is_hebrew(text):
    for ch in text:
        if '\u0590' <= ch <= '\u05FF':
            return True
    return False

def has_digits(text):
    return any(ch.isdigit() for ch in text)

def split_mixed_run(run):
    """פצל run שמכיל מעורב אנגלית+ספרות לרצף של runs"""
    text = run.text
    if not text:
        return
    has_latin = any('a' <= ch.lower() <= 'z' for ch in text)
    has_digit = any(ch.isdigit() for ch in text)
    if not (has_latin and has_digit and not is_hebrew(text)):
        return  # לא מעורב, אין צורך לפצל

    # בנה קבוצות של תווים לפי סוג
    import re
    groups = re.findall(r'[a-zA-Z]+|[0-9]+|[^a-zA-Z0-9]+', text)
    if len(groups) <= 1:
        return

    # קבל את ה-parent element
    parent = run._element.getparent()
    idx = list(parent).index(run._element)

    import copy
    # צור run חדש לכל קבוצה
    new_runs = []
    for group in groups:
        new_run_elem = copy.deepcopy(run._element)
        # עדכן טקסט
        t = new_run_elem.find(qn('w:t'))
        if t is None:
            from docx.oxml import OxmlElement
            t = OxmlElement('w:t')
            new_run_elem.append(t)
        t.text = group
        if group.startswith(' ') or group.endswith(' '):
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        new_runs.append((new_run_elem, group))

    # הסר את ה-run המקורי והכנס את החדשים
    run._element.text = ''
    for i, (new_run_elem, group) in enumerate(new_runs):
        parent.insert(idx + i, new_run_elem)
    parent.remove(run._element)

fix_run_font = None  # replaced by _fix_font_only

def _fix_font_only(run):
    """שנה רק פונט וגודל — מספרים תמיד David 13"""
    text = run.text
    if not text:
        return
    rPr = run._element.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        run._element.insert(0, rPr)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    # נקה את כל attributes של rFonts ותתחיל מחדש
    for attr in list(rFonts.attrib.keys()):
        del rFonts.attrib[attr]
    # עברית או ספרות בלבד — David 13
    # אנגלית (עם או בלי ספרות) — Times New Roman 11
    has_latin = any(('a' <= ch.lower() <= 'z') for ch in text)
    if is_hebrew(text) or (has_digits(text) and not has_latin):
        font, size = FONT_HEBREW, FONT_SIZE_HEBREW
    else:
        font, size = FONT_ENGLISH, FONT_SIZE_ENGLISH
    rFonts.set(qn('w:ascii'), font)
    rFonts.set(qn('w:hAnsi'), font)
    rFonts.set(qn('w:cs'), font)
    for tag in ['w:sz', 'w:szCs']:
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn('w:val'), str(size * 2))

def is_caption_text(text):
    s = text.strip()
    return s.startswith(("טבלה ", "שרטוט ", "תרשים "))

def process_document(input_bytes):
    # קרא את המסמך המקורי
    old_doc = Document(io.BytesIO(input_bytes))
    # צור מסמך חדש — זה נותן RTL מובנה
    new_doc = Document()

    # שוליים A4
    for section in new_doc.sections:
        section.page_width = PAGE_WIDTH
        section.page_height = PAGE_HEIGHT
        section.left_margin = MARGIN
        section.right_margin = MARGIN
        section.top_margin = MARGIN
        section.bottom_margin = MARGIN

    # הסר פסקה ריקה ראשונה של המסמך החדש
    for p in new_doc.paragraphs:
        p._element.getparent().remove(p._element)

    # העתק את כל ה-body מהמסמך הישן למסמך החדש
    old_body = old_doc.element.body
    new_body = new_doc.element.body
    sectPr = new_body.find(qn('w:sectPr'))

    for elem in list(old_body):
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        if tag == 'sectPr':
            continue
        elem_copy = copy.deepcopy(elem)
        if sectPr is not None:
            sectPr.addprevious(elem_copy)
        else:
            new_body.append(elem_copy)

    # עבור על כל הפסקאות ותקן פונט + רווח single
    for para in new_doc.paragraphs:
        # פצל runs מעורבים קודם
        for run in list(para.runs):
            split_mixed_run(run)
        # תקן פונט
        for run in para.runs:
            _fix_font_only(run)
        pPr = para._element.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            para._element.insert(0, pPr)
        spacing = pPr.find(qn('w:spacing'))
        if spacing is None:
            spacing = OxmlElement('w:spacing')
            pPr.append(spacing)
        spacing.set(qn('w:line'), '360')  # 1.5 בתוך פסקה
        spacing.set(qn('w:lineRule'), 'auto')
        spacing.set(qn('w:before'), '0')
        spacing.set(qn('w:after'), '0')  # אפס — שורות ריקות מהמקור מטפלות ברווח
        if is_caption_text(para.text.strip()):
            if pPr.find(qn('w:keepNext')) is None:
                kn = OxmlElement('w:keepNext')
                pPr.insert(0, kn)
            for run in para.runs:
                if run.text.strip():
                    run.bold = True
                    run.underline = True

    # תקן פסקאות ריקות — David 11 single
    for para in new_doc.paragraphs:
        if not para.text.strip() and not para._element.find('.//' + qn('w:drawing')):
            pPr = para._element.find(qn('w:pPr'))
            if pPr is None:
                pPr = OxmlElement('w:pPr')
                para._element.insert(0, pPr)
            spacing = pPr.find(qn('w:spacing'))
            if spacing is None:
                spacing = OxmlElement('w:spacing')
                pPr.append(spacing)
            spacing.set(qn('w:line'), '240')
            spacing.set(qn('w:lineRule'), 'auto')
            spacing.set(qn('w:before'), '0')
            spacing.set(qn('w:after'), '0')
            # פונט David 11
            rPr = pPr.find(qn('w:rPr'))
            if rPr is None:
                rPr = OxmlElement('w:rPr')
                pPr.append(rPr)
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.insert(0, rFonts)
            rFonts.set(qn('w:ascii'), FONT_HEBREW)
            rFonts.set(qn('w:hAnsi'), FONT_HEBREW)
            rFonts.set(qn('w:cs'), FONT_HEBREW)
            for tag in ['w:sz', 'w:szCs']:
                el = rPr.find(qn(tag))
                if el is None:
                    el = OxmlElement(tag)
                    rPr.append(el)
                el.set(qn('w:val'), str(11 * 2))

    # תקן פונט ב-rPr של pPr (משפיע על מספרי הרשימה)
    for para in new_doc.paragraphs:
        pPr = para._element.find(qn('w:pPr'))
        if pPr is not None:
            rPr = pPr.find(qn('w:rPr'))
            if rPr is None:
                rPr = OxmlElement('w:rPr')
                pPr.append(rPr)
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.insert(0, rFonts)
            for attr in list(rFonts.attrib.keys()):
                if 'hint' in attr or 'Theme' in attr or 'theme' in attr:
                    del rFonts.attrib[attr]
            rFonts.set(qn('w:ascii'), FONT_HEBREW)
            rFonts.set(qn('w:hAnsi'), FONT_HEBREW)
            rFonts.set(qn('w:cs'), FONT_HEBREW)
            for tag in ['w:sz', 'w:szCs']:
                el = rPr.find(qn(tag))
                if el is None:
                    el = OxmlElement(tag)
                    rPr.append(el)
                el.set(qn('w:val'), str(FONT_SIZE_HEBREW * 2))

    # נקה שורות ריקות כפולות — השאר רק אחת בין פסקאות
    body = new_doc.element.body
    prev_empty = False
    to_remove = []
    for elem in list(body):
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        if tag == 'p':
            try:
                from docx.text.paragraph import Paragraph
                para = Paragraph(elem, new_doc)
                text = para.text.strip()
                has_drawing = elem.find('.//' + qn('w:drawing')) is not None
                is_empty = not text and not has_drawing
                if is_empty and prev_empty:
                    to_remove.append(elem)
                prev_empty = is_empty
            except:
                prev_empty = False
        else:
            prev_empty = False
    for elem in to_remove:
        body.remove(elem)

    # תאי טבלאות — פונט + יישור מרכז
    for table in new_doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in list(para.runs):
                        split_mixed_run(run)
                    for run in para.runs:
                        _fix_font_only(run)
                    pPr = para._element.find(qn('w:pPr'))
                    if pPr is None:
                        pPr = OxmlElement('w:pPr')
                        para._element.insert(0, pPr)
                    jc = pPr.find(qn('w:jc'))
                    if jc is None:
                        jc = OxmlElement('w:jc')
                        pPr.append(jc)
                    jc.set(qn('w:val'), 'center')

    # עדכן פונט לכל ה-runs במסמך כולל תיבות טקסט
    from docx.oxml.ns import qn as qn2
    for elem in new_doc.element.body.iter():
        if elem.tag == qn2('w:r'):
            rPr = elem.find(qn2('w:rPr'))
            if rPr is not None:
                rFonts = rPr.find(qn2('w:rFonts'))
                if rFonts is not None:
                    # אם יש רק cs אבל חסר ascii — הוסף ascii ו-hAnsi
                    cs = rFonts.get(qn2('w:cs'))
                    ascii_f = rFonts.get(qn2('w:ascii'))
                    if cs and not ascii_f:
                        rFonts.set(qn2('w:ascii'), cs)
                        rFonts.set(qn2('w:hAnsi'), cs)
                    # נקה theme fonts
                    for attr in list(rFonts.attrib.keys()):
                        if 'Theme' in attr or 'theme' in attr:
                            del rFonts.attrib[attr]

    output = io.BytesIO()
    new_doc.save(output)
    output.seek(0)
    result = output.read()
    result = copy_media(input_bytes, result)
    return result

def copy_media(old_bytes, new_bytes):
    with zipfile.ZipFile(io.BytesIO(old_bytes)) as old_zip:
        old_rels = old_zip.read('word/_rels/document.xml.rels').decode('utf-8')
        img_rels = re.findall(r'<Relationship[^>]+image[^>]+/>', old_rels)
        media_files = {f: old_zip.read(f) for f in old_zip.namelist() if f.startswith('word/media/')}
        old_ct = old_zip.read('[Content_Types].xml').decode('utf-8')

    if not img_rels and not media_files:
        return new_bytes

    with zipfile.ZipFile(io.BytesIO(new_bytes)) as tmp:
        new_rels_text = tmp.read('word/_rels/document.xml.rels').decode('utf-8')
    existing_ids = set(re.findall(r'Id="(rId\d+)"', new_rels_text))

    id_map = {}
    counter = 100
    new_img_rels = []
    for rel in img_rels:
        while f'rId{counter}' in existing_ids:
            counter += 1
        old_id = re.search(r'Id="(rId\d+)"', rel).group(1)
        new_rel = re.sub(r'Id="rId\d+"', f'Id="rId{counter}"', rel)
        new_img_rels.append(new_rel)
        existing_ids.add(f'rId{counter}')
        id_map[old_id] = f'rId{counter}'
        counter += 1

    new_buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(new_bytes)) as zin:
        with zipfile.ZipFile(new_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            existing = set(zin.namelist())
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml' and id_map:
                    text = data.decode('utf-8')
                    for oid, nid in id_map.items():
                        text = text.replace(f'r:embed="{oid}"', f'r:embed="{nid}"')
                    data = text.encode('utf-8')
                elif item.filename == 'word/_rels/document.xml.rels' and new_img_rels:
                    insert = '\n  '.join(new_img_rels)
                    data = data.decode('utf-8').replace('</Relationships>', f'  {insert}\n</Relationships>').encode('utf-8')
                elif item.filename == '[Content_Types].xml' and media_files:
                    ct = data.decode('utf-8')
                    for t in re.findall(r'<Default[^>]+image[^>]+/>', old_ct):
                        ext = re.search(r'Extension="([^"]+)"', t)
                        if ext and ext.group(1) not in ct:
                            ct = ct.replace('</Types>', f'  {t}\n</Types>')
                    data = ct.encode('utf-8')
                zout.writestr(item, data)
            for fname, fdata in media_files.items():
                if fname not in existing:
                    zout.writestr(fname, fdata)
    new_buf.seek(0)
    return new_buf.read()

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
