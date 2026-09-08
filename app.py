import io
import re
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
import streamlit as st


# 1. ضبط اتجاه النص RTL داخل ملف Word
def set_rtl(paragraph_or_cell):
    pPr = paragraph_or_cell._element.get_or_add_pPr()
    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1")
    pPr.append(bidi)


# 2. بناء ملف Word المنسق
def build_smart_formatted_docx(raw_text):
    doc = docx.Document()

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

        # أول سطر: العنوان الرئيسي
        if first_line:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            set_rtl(p)
            clean_text = stripped.strip(" \"'()")
            run = p.add_run(clean_text)
            run.font.name = "Arial"
            run.font.size = Pt(20)
            run.bold = True
            run.font.color.rgb = RGBColor(31, 73, 125)  # أزرق داكن
            p.paragraph_format.space_after = Pt(14)
            first_line = False
            continue

        # العناوين الفرعية
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
            run.font.color.rgb = RGBColor(31, 73, 125)
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(6)

        # النقاط
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

        # النص العادي
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


# 3. إنشاء معاينة HTML تفاعلية بنفس شكل المستند
def generate_html_preview(raw_text):
    lines = raw_text.strip().split("\n")
    html_out = """
    <div style="
        background-color: #ffffff; 
        border: 1px solid #d1d5db; 
        border-radius: 8px; 
        padding: 35px; 
        font-family: Arial, sans-serif; 
        direction: rtl; 
        text-align: right; 
        box-shadow: 0px 4px 12px rgba(0,0,0,0.05);
        color: #222222;
        line-height: 1.8;
    ">
    """
    first_line = True

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        if first_line:
            clean_text = stripped.strip(" \"'()")
            html_out += f'<h1 style="color: #1F497D; font-size: 24px; font-weight: bold; margin-bottom: 16px; border-bottom: 2px solid #1F497D; padding-bottom: 8px;">{clean_text}</h1>'
            first_line = False
            continue

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
            html_out += f'<h2 style="color: #1F497D; font-size: 18px; font-weight: bold; margin-top: 20px; margin-bottom: 8px;">{stripped}</h2>'

        elif (
            stripped.startswith("-")
            or stripped.startswith("•")
            or stripped.startswith("إ")
            or ":" in stripped[:30]
        ):
            clean_bullet = re.sub(r"^[-•]\s*", "", stripped)
            if ":" in clean_bullet:
                parts = clean_bullet.split(":", 1)
                html_out += f'<p style="margin-right: 20px; margin-bottom: 6px; font-size: 15px;">• <strong style="color: #1F497D;">{parts[0]}:</strong>{parts[1]}</p>'
            else:
                html_out += f'<p style="margin-right: 20px; margin-bottom: 6px; font-size: 15px;">• {clean_bullet}</p>'

        else:
            html_out += f'<p style="font-size: 15px; margin-bottom: 10px; color: #333333;">{stripped}</p>'

    html_out += "</div>"
    return html_out


# --- الواجهة في Streamlit ---
st.set_page_config(
    page_title="مُنسّق ملخصات الكتب مع المعاينة", layout="wide"
)
st.title("📚 مُنسّق النصوص مع شاشة معاينة التنسيق")

# تقسيم الصفحة إلى عمودين
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("📝 إدخال النص")
    user_text = st.text_area(
        "ضع النص هنا:",
        height=450,
        placeholder="انسخ الملخص أو النص هنا...",
    )

with col2:
    st.subheader("👁️ معاينة التنسيق (Preview)")
    if user_text.strip():
        preview_html = generate_html_preview(user_text)
        st.markdown(preview_html, unsafe_allow_html=True)
    else:
        st.info("قم بكتابة أو لصق النص في الجهة اليسرى لرؤية المعاينة هنا.")

st.divider()

if user_text.strip():
    file_data = build_smart_formatted_docx(user_text)
    st.download_button(
        label="📥 تحميل ملف Word المنسق (.docx)",
        data=file_data,
        file_name="ملخص_منسق_جاهز.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
    )
