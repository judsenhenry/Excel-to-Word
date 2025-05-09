import streamlit as st
import tempfile
import os
import time
import re
import xlwings as xw
from pdf2image import convert_from_path
from PIL import Image
from docx import Document
from docx.shared import Cm, RGBColor
from docx.enum.text import WD_BREAK

# Find poppler
script_dir = os.path.dirname(os.path.abspath(__file__))
poppler_path = os.path.join(script_dir, "poppler", "Library", "bin")

st.set_page_config(page_title="Excel til Word Automation", layout="centered")
st.title("📄 Excel til Word Automation")

# Inputfelter
excel_file = st.file_uploader("Vælg Excel-fil", type=["xlsx", "xlsm"])
word_file = st.file_uploader("Vælg Word-fil", type=["docx"])
output_folder = st.text_input("Sti til outputmappe (brug absolut sti)")

top_crop = st.number_input("Beskær top (mm)", value=0)
bottom_crop = st.number_input("Beskær bund (mm)", value=0)
left_crop = st.number_input("Beskær venstre (mm)", value=0)
right_crop = st.number_input("Beskær højre (mm)", value=0)

sheet_selection = []

if excel_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(excel_file.read())
        tmp_path = tmp.name

    try:
        with xw.App(visible=False) as app:
            wb = app.books.open(tmp_path)
            sheet_names = [sheet.name for sheet in wb.sheets if sheet.visible]
            st.subheader("Vælg ark")
            sheet_selection = [name for name in sheet_names if st.checkbox(name)]
    except Exception as e:
        st.error(f"Kunne ikke indlæse ark: {e}")

def crop_image(image_path):
    img = Image.open(image_path)
    dpi = 300
    t = int(top_crop / 25.4 * dpi)
    b = int(bottom_crop / 25.4 * dpi)
    l = int(left_crop / 25.4 * dpi)
    r = int(right_crop / 25.4 * dpi)
    width, height = img.size
    cropped = img.crop((l, t, width - r, height - b))
    cropped.save(image_path)

def save_sheets_and_insert_images(tmp_excel_path):
    if not (tmp_excel_path and word_file and output_folder and sheet_selection):
        st.warning("Alle felter skal udfyldes.")
        return

    word_path = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    word_path.write(word_file.read())
    word_path.flush()

    with xw.App(visible=False) as app:
        wb = app.books.open(tmp_excel_path)
        total = len(sheet_selection)
        images_created = []

        progress = st.progress(0)
        status = st.empty()

        for i, sheet_name in enumerate(sheet_selection):
            sheet = wb.sheets[sheet_name]
            pdf_path = os.path.join(tempfile.gettempdir(), f"{sheet_name}.pdf")
            sheet.to_pdf(pdf_path)
            images = convert_from_path(pdf_path, dpi=300, poppler_path=poppler_path)
            for idx, img in enumerate(images, start=1):
                img_path = os.path.join(output_folder, f"{sheet_name}_{idx}.png")
                img.save(img_path)
                crop_image(img_path)
                images_created.append((sheet_name, img_path))
            progress.progress((i + 1) / total)
            status.text(f"Behandler: {sheet_name}")

    # Indsæt billeder i Word
    doc = Document(word_path.name)
    image_dict = {}
    pattern = re.compile(r"(.+?)_(\d+)\.png", re.IGNORECASE)
    for _, img_file in images_created:
        match = pattern.match(os.path.basename(img_file))
        if match:
            key, index = match.groups()
            image_dict.setdefault(key, []).append((int(index), img_file))

    for key in image_dict:
        image_dict[key].sort()

    for para in doc.paragraphs:
        for placeholder in list(image_dict.keys()):
            if f"{{{placeholder}}}" in para.text:
                para.text = para.text.replace(f"{{{placeholder}}}", "")
                run = para.add_run(f"{{{placeholder}}}")
                run.font.color.rgb = RGBColor(255, 255, 255)
                for i, (_, img_path) in enumerate(image_dict[placeholder]):
                    para.add_run().add_break()
                    para.add_run().add_picture(img_path, width=Cm(15))
                    if i < len(image_dict[placeholder]) - 1:
                        para.add_run().add_break(WD_BREAK.PAGE)

    output_docx = os.path.join(output_folder, "output_opdateret.docx")
    doc.save(output_docx)
    st.success(f"Dokument genereret: {output_docx}")

if st.button("Start proces"):
    with st.spinner("Kører..."):
        save_sheets_and_insert_images(tmp_path)
