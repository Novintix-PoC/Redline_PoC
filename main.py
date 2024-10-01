import pymupdf
import os
import pandas as pd
import streamlit as st
import re
from typing import List, Tuple
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTChar
import fontstyle

# Initialize variables
folder_path = ""
output_folder = ""
img_file = None
process_complete = False

st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;700&display=swap');
    
        .stApp {
            background: #0e4166;
            font-family: 'Roboto', sans-serif;
        }

        [data-testid=stHeader] {
            background: #0e4166; #for header
        }

        [data-testid=stFileUploaderDropzone],[data-baseweb=base-input] {
            background: rgb(0,48,73); #for header
            color:#f4a303
        }

        .stFileUploader span ,.stFileUploader small,[data-testid=stWidgetLabel] {
            color:#f4a303;
        }
        [data-testid="stBaseButton-secondary"]{
            color:#f4a303;
        }
        
        .title-container {
            background: rgb(0,48,73);
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
     
        
    <style>""", unsafe_allow_html=True)

# Main page with buttons
with st.container():
    col = st.columns((1, 3, 1))
    with col[0]:
        st.write("")
    with col[1]:
        st.markdown("""
        <div class="title-container">
            <h1 class="title">Redline Automation Tool</h1>
        </div>
        """, unsafe_allow_html=True)
    with col[2]:
        st.write("")


# User inputs
folder_path = st.text_input("Input Folder Path", folder_path)
excel_file = st.file_uploader("Choose an Excel file", type=["xlsx"])
img_file = st.file_uploader("Upload Image (if required)", type=["png", "jpg", "jpeg"])
output_folder = st.text_input("Output Folder Path", output_folder)

if st.button("Proceed"):
    if excel_file and folder_path and output_folder:
        try:
            excel_data = pd.read_excel(excel_file)
            excel_data = excel_data.applymap(lambda x: x.strip() if isinstance(x, str) else x)
            sorted_df = excel_data.sort_values(by=excel_data.columns[0])  # Sort by the first column
            grouped_data = excel_data.groupby('Part_Number')

            # Function definitions

            def overwrite(pdf_path, clean_copy, redline_copy):
                output_pdf = os.path.join(output_folder, f"{os.path.basename(pdf_path).replace('.pdf', '_task1.pdf')}")
                strike_out_and_replace(pdf_path, output_pdf, [(clean_copy, redline_copy)])
                return output_pdf

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
                        return text_info[i]['fontname'], text_info[i]['fontsize']

            def notes_addition(pdf_path, clean_copy, redline_copy):
                output_pdf = os.path.join(output_folder, f"{os.path.basename(pdf_path).replace('.pdf', '_task2.pdf')}")
                document = pymupdf.open(pdf_path)
                for page_num in range(len(document)):
                    page = document.load_page(page_num)
                    text_instances = page.search_for(clean_copy)
                    text_info = extract_text_with_font_info(pdf_path, clean_copy)
                    font, size = find_word_font_info(text_info, clean_copy)
                    styled_text = fontstyle.apply(redline_copy, font)
                    split_text = styled_text[:-3]  # Remove "[0m" error

                    if text_instances:
                        inst = text_instances[0]
                        page.insert_text((inst.x0 - 8, inst.y0), split_text, fontsize=size, color=(1, 0, 0), rotate=270)
                document.save(output_pdf)
                document.close()
                return output_pdf

            def cm_operation(pdf_path):
                output_pdf = os.path.join(output_folder, f"{os.path.basename(pdf_path).replace('.pdf', '_task3.pdf')}")
                if img_file:
                    img_path = os.path.join(output_folder, f"{os.path.basename(pdf_path).replace('.pdf', '_img')}" + os.path.splitext(img_file.name)[1])
                    with open(img_path, "wb") as img_temp:
                        img_temp.write(img_file.getbuffer())

                    image_width = 35
                    image_height = 30
                    left_margin = 515
                    top_margin = 700
                    bottom_left_x = left_margin
                    bottom_left_y = 792 - top_margin - image_height
                    top_right_x = left_margin + image_width
                    top_right_y = 792 - top_margin

                    document = pymupdf.open(pdf_path)
                    for page_num in range(len(document)):
                        page = document.load_page(page_num)
                        image_rect = pymupdf.Rect(bottom_left_x, bottom_left_y, top_right_x, top_right_y)
                        page.insert_image(image_rect, filename=img_path, rotate=270)

                    document.save(output_pdf)
                    document.close()
                    os.remove(img_path)
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
                            page.add_redact_annot(strikeout_rect, fill=(1, 0, 0))  # Strikeout in red
                            page.apply_redactions()
                            page.insert_text((strikeout_rect.x0 - 2, strikeout_rect.y0 - 3), original_text, fontsize=10, color=(0, 0, 0), rotate=270)
                            page.insert_text((strikeout_rect.x0 + 10, strikeout_rect.y0), replacement_text, fontsize=10, color=(1, 0, 0), rotate=270)
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
                            page.insert_text((strikeout_rect.x0 - 2, strikeout_rect.y0 - 3), original_text, fontsize=10, color=(0, 0, 0), rotate=270)
                            page.insert_text((strikeout_rect.x0 + 10, strikeout_rect.y0), replacement_text, fontsize=10, color=(1, 0, 0), rotate=270)
                document.save(output_path)
                document.close()

            # Processing grouped data
            for part_number, group in grouped_data:
                pdf_path = os.path.join(folder_path, f"{part_number}.pdf")
                if not os.path.exists(pdf_path):
                    st.error(f"PDF file for {part_number} not found.")
                    continue

                for _, row in group.iterrows():
                    task = row.get('Task')
                    clean_copy = row.get('Clean_Copy')
                    redline_copy = row.get('Redline_Copy')

                    if task == 'overwrite':
                        output_file = overwrite(pdf_path, clean_copy, redline_copy)
                        st.write(f"Task1 file saved at {output_file}")
                    elif task == 'notes addition':
                        output_file = notes_addition(pdf_path, clean_copy, redline_copy)
                        st.write(f"Task2 file saved at {output_file}")
                    elif task == 'cm operation':
                        output_file = cm_operation(pdf_path)
                        st.write(f"Task3 file saved at {output_file}")
                    elif task == 'rev replace':
                        output_file = os.path.join(output_folder, f"{os.path.basename(pdf_path).replace('.pdf', '_task4.pdf')}")
                        rev_replace(pdf_path, output_file)
                        st.write(f"Task4 file saved at {output_file}")

            process_complete = True
        except Exception as e:
            st.error(f"Error processing Excel data: {e}")

    else:
        st.warning("Please fill out all input fields before proceeding.")

# Completion message
if process_complete:
    st.success("Process completed successfully!")
