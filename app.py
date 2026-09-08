import io
import re
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
import streamlit as st

# اللون الأزرق المعتمد
BLUE_COLOR = RGBColor(31, 73, 125)  # #1F497D
BLACK_COLOR = RGBColor(0, 0, 0)


# 1. ضبط اتجاه النص RTL داخل ملف Word
def set_rtl(paragraph_or_cell):
    pPr = paragraph_or_cell._element.get_or_add_pPr()
    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1")
    pPr.append(bidi)


# 2. بناء ملف Word بالتنسيقات المحددة
def build_custom_docx(raw_text):
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

        # 1. العنوان الرئيسي: 16 Bold - أزرق
        if first_line:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            set_rtl(p)
            clean_text = stripped.strip(" \"'()")
            run = p.add_run(clean_text)
            run.font.name = "Arial"
            run.font.size = Pt(16)
            run.bold = True
            run.font.color.rgb = BLUE_COLOR
            p.paragraph_format.space_after = Pt(12)
            first_line = False
            continue

        # 2. العنوان الفرعي 1: 14 Bold على اليمين - أزرق
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
            run.font.color.rgb = BLUE_COLOR
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(4)

        # 3. العنوان الفرعي 2: 14 عادي - أسود (عند وجود نقطتين شارحة : في بداية السطر)
        elif ":" in stripped[:30] and not (
            stripped.startswith("-") or stripped.startswith("•")
        ):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            set_rtl(p)

            parts = stripped.split(":", 1)

            # جزء العنوان الفرعي 2
            run_sub2 = p.add_run(parts[0] + ":")
            run_sub2.font.name = "Arial"
            run_sub2.font.size = Pt(14)
            run_sub2.bold = False
            run_sub2.font.color.rgb = BLACK_COLOR

            # باقي محتوى الفقرة
            if len(parts) > 1:
                run_content = p.add_run(parts[1])
                run_content.font.name = "Arial"
                run_content.font.size = Pt(12)
                run_content.bold = False
                run_content.font.color.rgb = BLACK_COLOR

            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(4)

        # 4. المحتوى (النص العادي والنقاط): 12 عادي - أسود
        else:
            clean_text = re.sub(r"^[-•]\s*", "", stripped)
            is_bullet = stripped.startswith("-") or stripped.startswith("•")

            p = (
                doc.add_paragraph(style="List Bullet")
                if is_bullet
                else doc.add_paragraph()
            )
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            set_rtl(p)

            run = p.add_run(clean_text)
            run.font.name = "Arial"
            run.font.size = Pt(12)
            run.bold = False
            run.font.color.rgb = BLACK_COLOR
            p.paragraph_format.space_after = Pt(4)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# 3. إنشاء معاينة HTML بنفس القواعد
def generate_html_preview(raw_text):
    lines = raw_text.strip().split("\n")
    html_out = """
    <div style="
        background-color: #ffffff; 
        border: 1px solid #d1d5db; 
        border-radius: 8px; 
        padding: 30px; 
        font-family: Arial, sans-serif; 
        direction: rtl; 
        text-align: right; 
        box-shadow: 0px 4px 12px rgba(0,0,0,0.05);
        color: #000000;
        line-height: 1.8;
    ">
    """
    first_line = True

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        # 1. العنوان الرئيسي: 16 Bold أزرق
        if first_line:
            clean_text = stripped.strip(" \"'()")
            html_out += f'<h1 style="color: #1F497D; font-size: 21px; font-weight: bold; margin-bottom: 14px; border-bottom: 2px solid #1F497D; padding-bottom: 6px;">{clean_text}</h1>'
            first_line = False
            continue

        # 2. العنوان الفرعي 1: 14 Bold أزرق
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
            html_out += f'<h2 style="color: #1F497D; font-size: 18px; font-weight: bold; margin-top: 18px; margin-bottom: 6px; text-align: right;">{stripped}</h2>'

        # 3. العنوان الفرعي 2: 14 عادي أسود
        elif ":" in stripped[:30] and not (
            stripped.startswith("-") or stripped.startswith("•")
        ):
            parts = stripped.split(":", 1)
            html_out += f'<div style="margin-top: 10px; margin-bottom: 4px;"><span style="color: #000000; font-size: 18px; font-weight: normal;">{parts[0]}:</span><span style="color: #000000; font-size: 16px; font-weight: normal;">{parts[1] if len(parts)>1 else ""}</span></div>'

        # 4. المحتوى: 12 عادي أسود
        else:
            clean_text = re.sub(r"^[-•]\s*", "", stripped)
            is_bullet = stripped.startswith("-") or stripped.startswith("•")
            prefix = "• " if is_bullet else ""
            margin = "margin-right: 15px;" if is_bullet else ""
            html_out += f'<p style="font-size: 16px; color: #000000; margin-bottom: 6px; font-weight: normal; {margin}">{prefix}{clean_text}</p>'

    html_out += "</div>"
    return html_out


# --- الواجهة ---
st.set_page_config(
    page_title="مُنسّق الملخصات بالمقاسات الخاصة", layout="wide"
)
st.title("📚 مُنسّق النصوص مع المعاينة المباشرة")

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("📝 إدخال النص")
    user_text = st.text_area(
        "ضع النص هنا:",
        height=450,
        placeholder="انسخ الملخص أو النص هنا...",
    )

with col2:
    st.subheader("👁️ معاينة الشاشة والتنسيق")
    if user_text.strip():
        preview_html = generate_html_preview(user_text)
        st.markdown(preview_html, unsafe_allow_html=True)
    else:
        st.info("أدخل النص في الجهة اليسرى لرؤية التنسيق والمعاينة مباشرة.")

st.divider()

if user_text.strip():
    file_data = build_custom_docx(user_text)
    st.download_button(
        label="📥 تحميل ملف Word المنسق (.docx)",
        data=file_data,
        file_name="ملخص_منسق.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
    )
