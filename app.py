import streamlit as st
import fitz  # PyMuPDF
from PIL import Image
from io import BytesIO
from utils import create_zip, create_ppt, create_word

st.set_page_config(page_title="PDF to Carousel Converter", layout="wide")

st.title("📚 PDF to Carousel Converter (Poppler-Free)")

uploaded_files = st.file_uploader("Upload one or more PDF files", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    for uploaded_file in uploaded_files:
        st.header(f"File: {uploaded_file.name}")
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")

        pages = []
        for page in doc:
            zoom_x = 4.0  # X-axis zoom (4 times = ~400 DPI)
            zoom_y = 4.0  # Y-axis zoom
            matrix = fitz.Matrix(zoom_x, zoom_y)

            # Render page to an image with the zoom matrix (anti-aliasing effect)
            pix = page.get_pixmap(matrix=matrix, colorspace=fitz.csRGB, alpha=False)

            # Convert pixmap to image
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            pages.append(img)



        st.success(f"Converted {len(pages)} slides!")

        # HTML Carousel
        st.markdown("""
            <style>
            .carousel { display: flex; overflow-x: auto; gap: 20px; }
            .carousel img { border-radius: 10px; max-height: 500px; }
            </style>
        """, unsafe_allow_html=True)

        st.markdown('<div class="carousel">', unsafe_allow_html=True)
        for page in pages:
            img_byte_arr = BytesIO()
            page.save(img_byte_arr, format='PNG')
            st.image(img_byte_arr.getvalue(), use_column_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # Single Slide Download
        page_number = st.slider(f"Select a slide to download from {uploaded_file.name}", 1, len(pages), 1)
        img_byte_arr = BytesIO()
        pages[page_number - 1].save(img_byte_arr, format='PNG')

        st.download_button(
            label="📥 Download This Slide",
            data=img_byte_arr.getvalue(),
            file_name=f"{uploaded_file.name}_slide_{page_number}.png",
            mime="image/png"
        )

        # ZIP Download
        zip_buffer = create_zip(pages)
        st.download_button(
            label="📦 Download All Slides as ZIP",
            data=zip_buffer,
            file_name=f"{uploaded_file.name}_all_slides.zip",
            mime="application/zip"
        )

        # PowerPoint Download
        ppt_buffer = create_ppt(pages)
        st.download_button(
            label="📊 Download All Slides as PowerPoint",
            data=ppt_buffer,
            file_name=f"{uploaded_file.name}_slides.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

        # Word Download
        word_buffer = create_word(pages)
        st.download_button(
            label="📝 Download All Slides as Word Document",
            data=word_buffer,
            file_name=f"{uploaded_file.name}_slides.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        st.markdown("---")

st.markdown("Developed by **Irfan Ullah Khan** 🚀")
