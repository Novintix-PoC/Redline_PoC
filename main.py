#Latest working
import pymupdf
import os
import pandas as pd
import streamlit as st
import re
from typing import List, Tuple
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTChar
import fontstyle

# Initialize session state if not already done

if 'process_complete' not in st.session_state:
    st.session_state.process_complete = False  # To track processing status


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


# File Input Section
folder_path = st.text_input("Input Folder Path")
excel_file = st.file_uploader("Choose an Excel file", type=["xlsx"])
img_file = st.file_uploader("Upload Image (if required)", type=["png", "jpg", "jpeg"])
output_folder = st.text_input("Output Folder Path")

# Function to check for existing files
def check_files(folder, excel, output):
    if not folder or not os.path.exists(folder):
        st.error("The input folder path is invalid.")
        return False
    if not output or not os.path.exists(output):
        st.error("The output folder path is invalid.")
        return False
    if not excel:
        st.error("No Excel file uploaded.")
        return False
    return True

# Error Handling for Excel Read Operation
def read_excel_file(file):
    try:
        data = pd.read_excel(file)
        st.session_state.excel_data = data.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        return True
    except Exception as e:
        st.error(f"Failed to read the Excel file: {str(e)}")
        return False


if st.button("Proceed"):
    if check_files(folder_path, excel_file, output_folder) and read_excel_file(excel_file):
        try:
            excel_data = st.session_state.excel_data
            grouped_data = excel_data.groupby('Part_Number')

            # Function for Overwrite Operation
            def overwrite(pdf_path, clean_copy, redline_copy):
                output_pdf = os.path.join(output_folder, f"{os.path.basename(pdf_path).replace('.pdf', '_task1.pdf')}")
                strike_out_and_replace(pdf_path, output_pdf, [(clean_copy, redline_copy)])
                return output_pdf


            # To extract the font info (style and size) in the whole pdf
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

            # To find the font (style and size) for the Clean_copy in the PDF
            def find_word_font_info(text_info, search_word):
                
                word_length = len(search_word)
                for i in range(len(text_info) - word_length + 1):
                    # Check if the next few characters match the search word
                    if ''.join([text_info[j]['text'] for j in range(i, i + word_length)]) == search_word:
                        # Print font info for each character in the word
                        
                        for j in range(i, i + word_length):
                            text = text_info[j]['text']
                            font_name = text_info[j]['fontname']
                            font_size = text_info[j]['fontsize']
                            
                            return font_name , font_size

            # Function for Notes Addition Operation
            def notes_addition(pdf_path, clean_copy, redline_copy):
                output_pdf = os.path.join(output_folder, f"{os.path.basename(pdf_path).replace('.pdf', '_task2.pdf')}")
                document = pymupdf.open(pdf_path)

                for page_num in range(len(document)):
                    page = document.load_page(page_num)
                    text_instances = page.search_for(clean_copy)
                    text_info = extract_text_with_font_info(pdf_path, clean_copy) # Extract the font information
                    font , size = find_word_font_info(text_info, clean_copy) # Extracting the font style and font size
                    styled_text = fontstyle.apply(redline_copy, font) # Applying the font style and font size to the Redline_copy
                    split_text = styled_text[0 : len(styled_text) - 3] # to resolve the "[0m" error
                
                    if text_instances:
                        inst = text_instances[0]
                        notes_text = f"{redline_copy}"
                        page.insert_text((inst.x0 - 8, inst.y0), split_text, fontsize=size, color=(1, 0, 0), rotate=270)

                document.save(output_pdf)
                document.close()
                return output_pdf

            # Function for CM Operation
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
                    os.remove(img_path)  # Clean up the image file after processing
                    return output_pdf
                else:
                    st.error("No image uploaded for CM operation.")
                    return None

            # Strike out and replace function
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

            # Increment the revision number in the PDF
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
                            strikeout_annot = page.add_redact_annot(strikeout_rect, fill=(1, 0, 0))  # Strikeout in red
                            page.apply_redactions()
                            page.insert_text((strikeout_rect.x0 - 2, strikeout_rect.y0 - 3), original_text, fontsize=10, color=(0, 0, 0), rotate=270)
                            page.insert_text((strikeout_rect.x0 + 10, strikeout_rect.y0), replacement_text, fontsize=10, color=(1, 0, 0), rotate=270)

                document.save(output_path)
                document.close()

            for part_number, group in grouped_data:
                pdf_path = os.path.join(folder_path, f"{part_number}.pdf")
                
                if not os.path.exists(pdf_path):
                    st.warning(f"PDF file for {part_number} not found in input folder.")
                    continue

                # Process the PDF for each part
                intermediate_pdf = pdf_path
                for _, row in group.iterrows():
                    clean_copy = row['Clean_copy']
                    redline_copy = row['Redline_copy']
                    category = row['Category']

                    # Call appropriate functions for Overwrite, Notes, CM, etc.
                    if category == 'Overwrite':
                        intermediate_pdf = overwrite(intermediate_pdf, clean_copy, redline_copy)
                    elif category == 'Notes':
                        intermediate_pdf = notes_addition(intermediate_pdf, clean_copy, redline_copy)
                    elif category == 'CM':
                        intermediate_pdf = cm_operation(intermediate_pdf)
                        if intermediate_pdf is None:
                            break

                # Final processing after tasks (revision increment, cleanup)
                final_pdf_path = os.path.join(output_folder, f"{part_number}_redline.pdf")
                rev_replace(intermediate_pdf, final_pdf_path)
            

                # Increment the revision after processing
                temp_pdf_path = os.path.join(output_folder, f"temp_{part_number}.pdf")
                rev_replace(intermediate_pdf, temp_pdf_path)
                final_pdf_path = os.path.join(output_folder, f"{part_number}__redline.pdf")
                os.replace(temp_pdf_path, final_pdf_path)  # Replace temp with final

            # Cleanup: Remove all files except those ending with '_redline.pdf'
            for filename in os.listdir(output_folder):
                if not filename.endswith('_redline.pdf'):
                    file_path = os.path.join(output_folder, filename)
                    try:
                        os.remove(file_path)
                    except Exception as e:
                        st.warning(f"Could not delete file {file_path}: {e}")

            st.session_state.process_complete = True  # Mark process as complete

        except Exception as e:
            st.error(f"Error reading the Excel file: {e}")
    else:
        st.error("Please upload the Excel file and provide both input and output folder paths.")

# Display completion message if processing is complete
if st.session_state.process_complete:
    st.success("Processing complete!")

