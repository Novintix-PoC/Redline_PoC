import pymupdf
import os
import pandas as pd
import streamlit as st
import re
from typing import List, Tuple
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTChar
import fontstyle
import tempfile
import io
import zipfile

# Initialize session state if not already done
if 'excel_data' not in st.session_state:
    st.session_state.excel_data = None
if 'process_complete' not in st.session_state:
    st.session_state.process_complete = False  # To track processing status

st.markdown("""
    <style>
        [data-testid=stFileUploaderDropzone],[data-baseweb=base-input] {
            color:#f4a303
        }
        .title-container {
            backdrop-filter: blur(10px);
            border-radius: 30px;
            padding: 10px;
            box-shadow: 2px 4px 6px rgba(0, 0, 0, 0.1);
        }
        .title {
            color: #f4a303;
            font-size: 36px;
            font-weight: 700;
            text-align: center;
        }
        [data-testid="baseButton-secondary"],[data-testid="stBaseButton-secondary"] {
            margin-top: 10px;
            color: #f4a303;
        }

        
    </style>""", unsafe_allow_html=True)

# Main page with title
with st.container():
    col = st.columns((1, 3, 1))
    with col[1]:
        st.markdown("""
             <div class="title-container">
                <h1 class="title">Redline Automation Tool</h1>
            </div>
            """, unsafe_allow_html=True)

# File uploaders
excel_file = st.file_uploader("Upload the Input file", type=["xlsx"])
pdf_files = st.file_uploader("Upload Affected files", type=["pdf"], accept_multiple_files=True)
img_file = st.file_uploader("Upload Conformity Marking image (if required)", type=["png", "jpg", "jpeg"])

# Create a temporary directory for output
output_folder = tempfile.mkdtemp()

# Function definitions
def extract_text_with_font_info(pdf_path: str, search_word: str):
    text_info = []
    for page_layout in extract_pages(pdf_path):
        for element in page_layout:
            if isinstance(element, LTTextContainer):
                for text_line in element:
                    for character in text_line:
                        if isinstance(character, LTChar):
                            text_info.append({
                                "text": character.get_text(),
                                "fontname": character.fontname,
                                "fontsize": character.size
                            })
    return text_info

def find_word_font_info(text_info, search_word):
    word_length = len(search_word)
    for i in range(len(text_info) - word_length + 1):
        if ''.join([text_info[j]['text'] for j in range(i, i + word_length)]) == search_word:
            for j in range(i, i + word_length):
                text = text_info[j]['text']
                font_name = text_info[j]['fontname']
                font_size = text_info[j]['fontsize']
                return font_name, font_size

def overwrite(pdf_path, clean_copy, redline_copy):
    output_pdf = os.path.join(output_folder, f"{os.path.basename(pdf_path).replace('.pdf', '_task1.pdf')}")
    strike_out_and_replace(pdf_path, output_pdf, [(clean_copy, redline_copy)])
    return output_pdf

def notes_addition(pdf_path, clean_copy, redline_copy):
    output_pdf = os.path.join(output_folder, f"{os.path.basename(pdf_path).replace('.pdf', '_task2.pdf')}")
    document = pymupdf.open(pdf_path)
    for page_num in range(len(document)):
        page = document.load_page(page_num)
        text_instances = page.search_for(clean_copy)
        text_info = extract_text_with_font_info(pdf_path, clean_copy)
        font, size = find_word_font_info(text_info, clean_copy)
        styled_text = fontstyle.apply(redline_copy, font)
        split_text = styled_text[0: len(styled_text) - 3]
        if text_instances:
            inst = text_instances[0]
            page.insert_text((inst.x0 - 8, inst.y0), split_text, fontsize=size, color=(1, 0, 0), rotate=270)
    document.save(output_pdf)
    document.close()
    return output_pdf

def cm_operation(pdf_path):
    output_pdf = os.path.join(output_folder, f"{os.path.basename(pdf_path).replace('.pdf', '_task3.pdf')}")
    if img_file:
        img_data = img_file.getvalue()
        image_width, image_height = 35, 30
        left_margin, top_margin = 515, 700
        bottom_left_x, bottom_left_y = left_margin, 792 - top_margin - image_height
        top_right_x, top_right_y = left_margin + image_width, 792 - top_margin
        document = pymupdf.open(pdf_path)
        for page_num in range(len(document)):
            page = document.load_page(page_num)
            image_rect = pymupdf.Rect(bottom_left_x, bottom_left_y, top_right_x, top_right_y)
            page.insert_image(image_rect, stream=img_data, rotate=270)
        document.save(output_pdf)
        document.close()
        return output_pdf
    else:
        st.error("No image uploaded for CM operation.")
        return None

def strike_out_and_replace(pdf_path: str, output_path: str, replacements: List[Tuple[str, str]]):
    document = pymupdf.open(pdf_path)
    for page_num in range(len(document)):
        page = document.load_page(page_num)
        for original_text, replacement_text in replacements:
            text_instances = page.search_for(original_text)
            for inst in text_instances:
                strikeout_rect = pymupdf.Rect(
                    inst.x0 + (inst.x1 - inst.x0) / 2 - 0.5,
                    inst.y0,
                    inst.x0 + (inst.x1 - inst.x0) / 2 + 0.5,
                    inst.y1
                )
                page.add_redact_annot(strikeout_rect, fill=(1, 0, 0))
                page.apply_redactions()
                page.insert_text((strikeout_rect.x0 - 2, strikeout_rect.y0 - 3),
                                 original_text, fontsize=10, color=(0, 0, 0), rotate=270)
                page.insert_text((strikeout_rect.x0 + 10, strikeout_rect.y0),
                                 replacement_text, fontsize=10, color=(1, 0, 0), rotate=270)
    document.save(output_path)
    document.close()

def rev_replace(pdf_path: str, output_path: str):
    document = pymupdf.open(pdf_path)
    pattern = re.compile(r'rev(\d+)')
    for page_num in range(len(document)):
        page = document.load_page(page_num)
        page_text = page.get_text("text")
        matches = pattern.finditer(page_text)
        for match in matches:
            original_text = match.group(0)
            number = int(match.group(1))
            replacement_text = f"rev{number + 1:02}"
            text_instances = page.search_for(original_text)
            for inst in text_instances:
                strikeout_rect = pymupdf.Rect(
                    inst.x0 + (inst.x1 - inst.x0) / 2 - 0.5,
                    inst.y0,
                    inst.x0 + (inst.x1 - inst.x0) / 2 + 0.5,
                    inst.y1
                )
                page.add_redact_annot(strikeout_rect, fill=(1, 0, 0))
                page.apply_redactions()
                page.insert_text((strikeout_rect.x0 - 2, strikeout_rect.y0 - 3),
                                 original_text, fontsize=10, color=(0, 0, 0), rotate=270)
                page.insert_text((strikeout_rect.x0 + 10, strikeout_rect.y0),
                                 replacement_text, fontsize=10, color=(1, 0, 0), rotate=270)
    document.save(output_path)
    document.close()

if st.button("Proceed"):
    if excel_file and pdf_files:
        with st.spinner("Processing..."):
            try:
                # Process Excel file
                excel_data = pd.read_excel(excel_file)
                excel_data = excel_data.applymap(lambda x: x.strip() if isinstance(x, str) else x)
                sorted_df = excel_data.sort_values(by=excel_data.columns[0])
                grouped_data = excel_data.groupby('Part_Number')

                # Process PDF files
                for pdf_file in pdf_files:
                    pdf_path = os.path.join(output_folder, pdf_file.name)
                    with open(pdf_path, "wb") as f:
                        f.write(pdf_file.getvalue())
                    
                    part_number = os.path.splitext(pdf_file.name)[0]
                    if part_number in grouped_data.groups:
                        group = grouped_data.get_group(part_number)
                        intermediate_pdf = pdf_path
                        for index, row in group.iterrows():
                            clean_copy = row['Clean_copy']
                            redline_copy = row['Redline_copy']
                            category = row['Category']

                            if category == 'Overwrite':
                                intermediate_pdf = overwrite(intermediate_pdf, clean_copy, redline_copy)
                            elif category == 'Notes':
                                intermediate_pdf = notes_addition(intermediate_pdf, clean_copy, redline_copy)
                            elif category == 'CM':#'Conformity_Marking':
                                intermediate_pdf = cm_operation(intermediate_pdf)
                                if intermediate_pdf is None:
                                    break
                            else:
                                st.warning(f"Unknown category '{category}' for part {part_number}. No action taken.")

                        # Increment the revision after processing
                        temp_pdf_path = os.path.join(output_folder, f"temp_{part_number}.pdf")
                        rev_replace(intermediate_pdf, temp_pdf_path)
                        final_pdf_path = os.path.join(output_folder, f"{part_number}__redline.pdf")
                        os.replace(temp_pdf_path, final_pdf_path)
                    else:
                        st.warning(f"No data found in Excel for PDF: {pdf_file.name}")

                # Cleanup: Remove all files except those ending with '_redline.pdf'
                for filename in os.listdir(output_folder):
                    if not filename.endswith('_redline.pdf'):
                        file_path = os.path.join(output_folder, filename)
                        try:
                            os.remove(file_path)
                        except Exception as e:
                            st.warning(f"Could not delete file {file_path}: {e}")

                st.session_state.process_complete = True

            except Exception as e:
                st.error(f"An error occurred: {str(e)}")
    else:
        st.error("Please upload the Excel file and PDF files.")

# Display completion message and provide a download button for all files
if st.session_state.process_complete:
    st.success("Processing complete!")
    
    # Create a zip file containing all processed PDFs
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for filename in os.listdir(output_folder):
            if filename.endswith('_redline.pdf'):
                file_path = os.path.join(output_folder, filename)
                zip_file.write(file_path, filename)
    
    # Offer the zip file for download
    st.download_button(
        label="Download All the processed PDFs",
        data=zip_buffer.getvalue(),
        file_name="Output.zip",
        mime="application/zip"
    )
