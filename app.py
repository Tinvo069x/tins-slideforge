import os
import streamlit as st
from docx import Document
from pptx import Presentation
from pptx.util import Pt

def convert_docx_to_pptx(docx_file):
    # Đọc file Word
    doc = Document(docx_file)

    prs = Presentation()
    title_slide_layout = prs.slide_layouts[0]
    content_slide_layout = prs.slide_layouts[1]

    # Slide 1: Title
    slide = prs.slides.add_slide(title_slide_layout)
    slide.shapes.title.text = doc.paragraphs[0].text.strip()
    slide.placeholders[1].text = "Summary generated from Word"

    current_slide = None
    text_box = None
    content_lines = []

    for para in doc.paragraphs[1:]:
        text = para.text.strip()
        if not text:
            continue

        if text.endswith(":"):  # New slide
            if current_slide and text_box and content_lines:
                text_box.text = "\n".join(content_lines)
                for p in text_box.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(20)
                content_lines = []

            current_slide = prs.slides.add_slide(content_slide_layout)
            current_slide.shapes.title.text = text.replace(":", "")
            text_box = current_slide.placeholders[1].text_frame
            text_box.clear()
        else:
            if current_slide:
                content_lines.append(text)

    # Dump last content
    if current_slide and text_box and content_lines:
        text_box.text = "\n".join(content_lines)
        for p in text_box.paragraphs:
            for run in p.runs:
                run.font.size = Pt(20)

    # Xuất ra file pptx trong bộ nhớ
    output_path = "converted.pptx"
    prs.save(output_path)
    return output_path

# ================= Streamlit UI =================
st.set_page_config(page_title="Tins_SlideForge", page_icon="📑", layout="centered")
st.title("📑 Tins_SlideForge - Word → PowerPoint")

uploaded_file = st.file_uploader("📂 Upload Word file (.docx)", type=["docx"])

if uploaded_file:
    with st.spinner("⏳ Converting..."):
        output_pptx = convert_docx_to_pptx(uploaded_file)

    st.success("✅ Conversion completed!")
    with open(output_pptx, "rb") as f:
        st.download_button(
            "⬇️ Download PowerPoint",
            f,
            file_name="Converted.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
