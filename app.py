import io
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
import streamlit as st


def set_rtl(paragraph_or_cell):
    pPr = paragraph_or_cell._element.get_or_add_pPr()
    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1")
    pPr.append(bidi)


def build_docx(h1_text, h2_text, body_text, table_data_raw):
    doc = docx.Document()

    if h1_text.strip():
        p1 = doc.add_paragraph()
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_rtl(p1)
        run1 = p1.add_run(h1_text)
        run1.font.name = "Arial"
        run1.font.size = Pt(16)
        run1.bold = True

    if h2_text.strip():
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_rtl(p2)
        run2 = p2.add_run(h2_text)
        run2.font.name = "Arial"
        run2.font.size = Pt(14)
        run2.bold = True

    if body_text.strip():
        for line in body_text.split("\n"):
            if line.strip():
                pb = doc.add_paragraph()
                pb.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                set_rtl(pb)
                run_b = pb.add_run(line)
                run_b.font.name = "Arial"
                run_b.font.size = Pt(12)
                run_b.bold = False

    if table_data_raw.strip():
        rows = [
            row.split("\t")
            for row in table_data_raw.strip().split("\n")
            if row
        ]
        if rows:
            max_cols = max(len(r) for r in rows)
            table = doc.add_table(rows=len(rows), cols=max_cols)
            table.alignment = WD_TABLE_ALIGNMENT.RIGHT

            tblPr = table._element.xpath("w:tblPr")
            if tblPr:
                bidiVisual = OxmlElement("w:bidiVisual")
                tblPr[0].append(bidiVisual)

            for r_idx, row in enumerate(rows):
                for c_idx, cell_value in enumerate(row):
                    cell = table.cell(r_idx, c_idx)
                    cell.text = cell_value.strip()

                    for p in cell.paragraphs:
                        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        set_rtl(p)
                        for run in p.runs:
                            run.font.name = "Arial"
                            run.font.size = Pt(12)
                            if r_idx == 0:
                                run.bold = True

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


st.set_page_config(page_title="أداة تنسيق المستندات", layout="centered")
st.title("📄 أداة تنسيق النصوص وتحويلها إلى Word")

h1 = st.text_input("العنوان الرئيسي (حجم 16 - Bold):")
h2 = st.text_input("العنوان الفرعي (حجم 14 - Bold):")
body = st.text_area("النص العادي (حجم 12):", height=200)

st.write("### الجدول (اختياري)")
st.caption("انسخ الجدول من إكسل أو نوت باد والصقه هنا (مفصول بـ Tab):")
table_raw = st.text_area("بيانات الجدول:", height=120)

if st.button("تنسيق وإنشاء الملف"):
    doc_buffer = build_docx(h1, h2, body, table_raw)
    st.download_button(
        label="📥 تحميل ملف Word المتنسق",
        data=doc_buffer,
        file_name="Formatted_Document.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    st.success("تم التنسيق بنجاح!")
