import io
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
import streamlit as st


# ضبط اتجاه RTL للفقرات والخلايا
def set_rtl(paragraph_or_cell):
    pPr = paragraph_or_cell._element.get_or_add_pPr()
    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1")
    pPr.append(bidi)


# محرك المعالجة التلقائية للنص الكامل
def process_full_text_to_docx(raw_text):
    doc = docx.Document()
    lines = raw_text.strip().split("\n")

    is_first_line = True
    in_table_mode = False
    table_buffer = []

    for line in lines:
        stripped = line.strip()

        # 1. التعرّف على الجداول (الأسطر اللي فيها علامة | أو Tabs)
        if "|" in stripped or "\t" in stripped:
            in_table_mode = True
            # تنظيف الصف وتقطيعه
            if "|" in stripped:
                row_data = [
                    c.strip() for c in stripped.split("|") if c.strip() != ""
                ]
            else:
                row_data = [c.strip() for c in stripped.split("\t") if c.strip()]

            # نمرر أسطر التنسيق مثل |-|-|
            if row_data and not all(set(c) <= set("-:") for c in row_data):
                table_buffer.append(row_data)
            continue

        # لو خرجنا من الجدول، نرسمه في الملف
        if in_table_mode and table_buffer:
            create_rtl_table(doc, table_buffer)
            table_buffer = []
            in_table_mode = False

        if not stripped:
            continue

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_rtl(p)
        run = p.add_run(stripped)
        run.font.name = "Arial"

        # 2. العنوان الرئيسي (أول سطر في النص)
        if is_first_line:
            run.font.size = Pt(16)
            run.bold = True
            is_first_line = False

        # 3. العناوين الفرعية (الأسطر التي تبدأ بـ # أو * أو أسطر قصيرة تنتهي بـ :)
        elif (
            stripped.startswith("#")
            or stripped.startswith("* ")
            or (len(stripped) < 40 and stripped.endswith(":"))
        ):
            clean_text = stripped.lstrip("#* ").strip()
            p.text = ""  # إعادة تفريغ الفقرة لكتابة النص المنظف
            run_h2 = p.add_run(clean_text)
            run_h2.font.name = "Arial"
            run_h2.font.size = Pt(14)
            run_h2.bold = True

        # 4. باقي الكلام العادي
        else:
            run.font.size = Pt(12)
            run.bold = False

    # طباعة الجدول في حال كان آخر النص
    if table_buffer:
        create_rtl_table(doc, table_buffer)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# دالة رسم الجداول RTL
def create_rtl_table(doc, table_data):
    if not table_data:
        return
    max_cols = max(len(r) for r in table_data)
    table = doc.add_table(rows=len(table_data), cols=max_cols)
    table.alignment = WD_TABLE_ALIGNMENT.RIGHT

    tblPr = table._element.xpath("w:tblPr")
    if tblPr:
        bidiVisual = OxmlElement("w:bidiVisual")
        tblPr[0].append(bidiVisual)

    for r_idx, row in enumerate(table_data):
        for c_idx, cell_value in enumerate(row):
            if c_idx < max_cols:
                cell = table.cell(r_idx, c_idx)
                cell.text = cell_value

                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    set_rtl(p)
                    for run in p.runs:
                        run.font.name = "Arial"
                        run.font.size = Pt(12)
                        if r_idx == 0:  # هيدر الجدول Bold
                            run.bold = True


# --- واجهة المستخدم الذكية ---
st.set_page_config(page_title="المُنسّق التلقائي الشامل", layout="centered")
st.title("⚡ المُنسّق التلقائي للمستندات (Word Auto-Formatter)")
st.write(
    "ارزع نصك الكامل هنا مباشرة (سواء من ChatGPT أو نوت باد أو إكسل)، والأداة هتفصله وتنسقه أوتوماتيك!"
)

full_input = st.text_area(
    "ضع النص الكامل هنا:",
    height=350,
    placeholder="""أول سطر هيكون هو العنوان الرئيسي (16 Bold) تلقائياً...

عنوان فرعي: (أي سطر قصير بـ : أو يبدأ بـ # هيكون 14 Bold)
ده هيكون كلام عادي بحجم 12 من غير أي تدخل منك.

| جدول | تلقائي | RTL |
|--- |--- |--- |
| صف 1 | بيانات | 12 |""",
)

if st.button("تنسيق وتحويل النص أوتوماتيكياً 🚀"):
    if full_input.strip():
        doc_file = process_full_text_to_docx(full_input)
        st.download_button(
            label="📥 تحميل ملف Word المتنسق جاهز",
            data=doc_file,
            file_name="Formatted_Document.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        st.success("تم التنسيق التلقائي بنجاح!")
    else:
        st.error("يرجى إدخال نص أولاً!")
