import io
import re
import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor
import streamlit as st


def set_rtl(paragraph_or_cell):
    pPr = paragraph_or_cell._element.get_or_add_pPr()
    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1")
    pPr.append(bidi)


def build_smart_formatted_docx(raw_text):
    doc = docx.Document()

    # الهوامش
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    lines = raw_text.strip().split("\n")
    first_line = True

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        # 1. أول سطر يتنسق كـ "عنوان رئيسي كبير"
        if first_line:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            set_rtl(p)
            clean_text = stripped.strip(" \"'()")
            run = p.add_run(clean_text)
            run.font.name = "Arial"
            run.font.size = Pt(20)
            run.bold = True
            run.font.color.rgb = RGBColor(31, 73, 125)  # الأزرق الداكن
            p.paragraph_format.space_after = Pt(14)
            first_line = False
            continue

        # 2. كشف العناوين الفرعية (التي تحتوي على أرقام صفحات أو عناوين أقسام)
        if (
            "ص " in stripped
            or "ص1" in stripped
            or "ص 9" in stripped
            or (len(stripped) < 60 and not stripped.startswith("•"))
        ) and (
            ":" not in stripped[:20]
            and not stripped.startswith("-")
            and not stripped.startswith("•")
        ):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            set_rtl(p)
            run = p.add_run(stripped)
            run.font.name = "Arial"
            run.font.size = Pt(14)
            run.bold = True
            run.font.color.rgb = RGBColor(31, 73, 125)  # الأزرق الداكن
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(6)

        # 3. كشف النقاط
        elif (
            stripped.startswith("-")
            or stripped.startswith("•")
            or stripped.startswith("إ")
            or ":" in stripped[:30]
        ):
            clean_bullet = re.sub(r"^[-•]\s*", "", stripped)
            p = doc.add_paragraph(style="List Bullet")
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            set_rtl(p)

            # جعل الجزء قبل النقطتين Bold لو موجود
            if ":" in clean_bullet:
                parts = clean_bullet.split(":", 1)
                r1 = p.add_run(parts[0] + ":")
                r1.font.name = "Arial"
                r1.font.size = Pt(11)
                r1.bold = True
                r1.font.color.rgb = RGBColor(31, 73, 125)

                r2 = p.add_run(parts[1])
                r2.font.name = "Arial"
                r2.font.size = Pt(11)
            else:
                r = p.add_run(clean_bullet)
                r.font.name = "Arial"
                r.font.size = Pt(11)

            p.paragraph_format.space_after = Pt(4)

        # 4. باقي الفقرات العادية
        else:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            set_rtl(p)
            run = p.add_run(stripped)
            run.font.name = "Arial"
            run.font.size = Pt(11)
            p.paragraph_format.space_after = Pt(6)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# الواجهة
st.set_page_config(page_title="مُنسّق الكتب والملخصات", layout="centered")
st.title("📚 مُنسّق ملخصات الكتب التلقائي")

user_text = st.text_area("ضع نص الملخص هنا:", height=350)

if st.button("تنسيق وتحويل إلى Word 🚀"):
    if user_text.strip():
        file_data = build_smart_formatted_docx(user_text)
        st.download_button(
            label="📥 تحميل الملف المنسق جاهز",
            data=file_data,
            file_name="ملخص_منسق.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        st.success("تم التنسيق بنجاح!")
