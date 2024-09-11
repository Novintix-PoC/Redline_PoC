# pip install pymupdf pandas streamlit fontstyle pdfminer.six

import pymupdf
import re
import os
import pandas as pd
import streamlit as st
from typing import List, Tuple
import streamlit as st
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTChar
import fontstyle
# Function to switch pages
def switch_page(page_name):
    st.session_state.current_page = page_name

# Initialize session state if not already done
if 'current_page' not in st.session_state:
    st.session_state.current_page = 'main'

# Main page with buttons
if st.session_state.current_page == 'main':
    with st.container():
        col = st.columns((1, 3, 1))
        with col[0]:
            st.write("")
        with col[1]:
            st.title("Redline Automation")
        with col[2]:
            st.write("")

    st.write("please select the type of change")
    columns = st.columns((1, 1, 1))
    with columns[0]:
        
        if st.button("Overwrite", use_container_width=True):
            switch_page('Overwrite')

    with columns[1]:
        if st.button("Conformity Marking", use_container_width=True):
            switch_page('conformity Marking')

    with columns[2]:
        if st.button("Notes Addition", use_container_width=True):
            switch_page('Notes Addition')

# Page 1
elif st.session_state.current_page == 'Overwrite':
    
    def find_and_replace(pdf_path: str, output_path: str, replacements: List[Tuple[str, str]]):
        document = pymupdf.open(pdf_path)
        pattern = re.compile(r'rev(\d+)')

        for page_num in range(len(document)):
            page = document.load_page(page_num)
            page_text = page.get_text("text")
            matches = pattern.finditer(page_text) 

            for match in matches:
                original_text = match.group(0)
                number = int(match.group(1))
                replacement_text = f"rev{number+1:02}"
                text_instances = page.search_for(original_text)

                for inst in text_instances:
                    strikeout_rect = pymupdf.Rect(
                        inst.x0 + (inst.x1 - inst.x0) / 2 - 0.5,
                        inst.y0,
                        inst.x0 + (inst.x1 - inst.x0) / 2 + 0.5,
                        inst.y1
                    )
                    strikeout_annot = page.add_redact_annot(strikeout_rect, fill=(1, 0, 0))
                    page.apply_redactions()
                    page.insert_text((strikeout_rect.x0-2, strikeout_rect.y0-3), original_text, fontsize=10, color=(0, 0, 0), rotate=270)
                    page.insert_text((strikeout_rect.x0 + 10, strikeout_rect.y0), replacement_text, fontsize=12, color=(1, 0, 0), rotate=270)
                    
        document.save(output_path)
        document.close()

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
                    strikeout_annot = page.add_redact_annot(strikeout_rect, fill=(1, 0, 0))
                    page.apply_redactions()
                    page.insert_text((strikeout_rect.x0-2, strikeout_rect.y0-3), original_text, fontsize=10, color=(0, 0, 0), rotate=270)
                    page.insert_text((strikeout_rect.x0 + 10, strikeout_rect.y0), replacement_text, fontsize=12, color=(1, 0, 0), rotate=270)

        document.save(output_path)
        document.close()

    def process_excel_and_replace(excel_file: pd.DataFrame, folder_path: str, output_folder: str):
        df = excel_file
        
        for index, row in df.iterrows():
            part_number = row['Part_Number']
            banner_copy = row['Clean_copy']
            redline_copy = row['Redline_copy']
            
            pdf_path = os.path.join(folder_path, f"{part_number}.pdf")
            intermediate_pdf = os.path.join(output_folder, f"{part_number}_intermediate.pdf")
            output_pdf = os.path.join(output_folder, f"{part_number}_Redline.pdf")
            
            find_and_replace(pdf_path, intermediate_pdf, [(banner_copy, redline_copy)])
            strike_out_and_replace(intermediate_pdf, output_pdf, [(banner_copy, redline_copy)])
            os.remove(intermediate_pdf)

    # Streamlit interface
    col = st.columns((1, 6, 1))
    with col[0]:
        st.write("")
    with col[1]:
        st.title("Redline Automation Tool")
    with col[2]:
        st.write("")

    st.write("Please provide the necessary inputs to process your PDF files.")

    folder_path = st.text_input("Enter the folder path where the PDF files are located:")
    excel_file = st.file_uploader("Upload the Excel file:", type=["xlsx"])
    output_folder = st.text_input("Enter the folder path where the output PDF files should be saved:")

    if st.button("Start Process"):
        if folder_path and excel_file and output_folder:
            excel_data = pd.read_excel(excel_file)
            process_excel_and_replace(excel_data, folder_path, output_folder)
            st.success("Process completed successfully!")
        else:
            st.error("Please provide all inputs.")

    if st.button("Go Back"):
        switch_page('main')
# Page 2
elif st.session_state.current_page == 'conformity Marking':
    col = st.columns((1, 6, 1))
    with col[0]:
        st.write("")
    with col[1]:
        st.title("Redline Automation Tool")
    with col[2]:
        st.write("")

    st.write("Please provide the necessary inputs to process your PDF files.")

    input_folder = st.text_input("Enter the folder path where your PDF files reside:")
    img_file = st.file_uploader("Upload the Image:", type=["png", "jpg", "jpeg"])
    output_folder = st.text_input("Enter the folder path where the output PDF files should be saved:")

    if st.button("Start Process"):

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
                    replacement_text = f"rev{number+1:02}"
                    text_instances = page.search_for(original_text)

                    for inst in text_instances:
                        strikeout_rect = pymupdf.Rect(
                            inst.x0 + (inst.x1 - inst.x0) / 2 - 0.5,
                            inst.y0,
                            inst.x0 + (inst.x1 - inst.x0) / 2 + 0.5,
                            inst.y1
                        )
                        strikeout_annot = page.add_redact_annot(strikeout_rect, fill=(1, 0, 0))
                        page.apply_redactions()
                        page.insert_text((strikeout_rect.x0-2, strikeout_rect.y0-3), original_text, fontsize=10, color=(0, 0, 0), rotate=270)
                        page.insert_text((strikeout_rect.x0 + 10, strikeout_rect.y0), replacement_text, fontsize=12, color=(1, 0, 0), rotate=270)
                        
            document.save(output_path)
            document.close()

        if input_folder and img_file and output_folder:
            try:
                img_path = os.path.join(output_folder, "temp_img" + os.path.splitext(img_file.name)[1])
                
                # Save uploaded image to a temporary file
                with open(img_path, "wb") as img_temp:
                    img_temp.write(img_file.getbuffer())

                # Check if input and output directories exist
                if not os.path.exists(input_folder):
                    st.error("Input folder does not exist.")
                if not os.path.exists(output_folder):
                    os.makedirs(output_folder)

                # Loop through all PDF files in the input folder
                for filename in os.listdir(input_folder):
                    if filename.endswith(".pdf"):
                        # Create full path for input and output PDFs
                        input_file = os.path.join(input_folder, filename)
                        intermediate_filename = os.path.splitext(filename)[0] + "_Intermediate.pdf"
                        intermediate_pdf = os.path.join(output_folder, intermediate_filename)
                        output_filename = os.path.splitext(filename)[0] + "_Redline.pdf"
                        output_pdf = os.path.join(output_folder, output_filename)

                        # Open the PDF document
                        image_width = 35
                        image_height = 30
                        left_margin = 515
                        top_margin = 700

                        bottom_left_x = left_margin
                        bottom_left_y = 792 - top_margin - image_height
                        top_right_x = left_margin + image_width
                        top_right_y = 792 - top_margin

                        image_rectangle = pymupdf.Rect(bottom_left_x, bottom_left_y, top_right_x, top_right_y)

                        # Open the PDF document
                        file_handle = pymupdf.open(input_file)
                        first_page = file_handle[0]

                        # Add the image to the first page
                        try:
                            first_page.insert_image(image_rectangle, filename=img_path, rotate=270)
                            # Save the intermediate PDF
                            file_handle.save(intermediate_pdf)
                            
                            # Perform redline replacement on the intermediate file and save the final output
                            rev_replace(intermediate_pdf, output_pdf)
                            
                        except Exception as e:
                            st.error(f"An error occurred while processing {filename}: {e}")
                        finally:
                            # Remove the intermediate file
                            if os.path.exists(intermediate_pdf):
                                os.remove(intermediate_pdf)
                
                # Remove the temporary image file
                os.remove(img_path)

                st.success("All PDF files in the input folder have been processed!")
            except Exception as e:
                st.error(f"An error occurred: {e}")
        else:
            st.error("Please provide all inputs.")

    if st.button("Go Back"):
        switch_page('main')

# Page 3
elif st.session_state.current_page == 'Notes Addition':
    # To replace the (revision) number 
    def rev_replace(pdf_path: str, output_path: str, replacements: List[Tuple[str, str]]):
        document = pymupdf.open(pdf_path)
        pattern = re.compile(r'rev(\d+)')

        for page_num in range(len(document)):
            page = document.load_page(page_num)
            page_text = page.get_text("text")
            matches = pattern.finditer(page_text)

            for match in matches:
                original_text = match.group(0)
                number = int(match.group(1))
                replacement_text = f"rev{number+1:02}"
                text_instances = page.search_for(original_text)

                for inst in text_instances:
                    strikeout_rect = pymupdf.Rect(
                        inst.x0 + (inst.x1 - inst.x0) / 2 - 0.5,
                        inst.y0,
                        inst.x0 + (inst.x1 - inst.x0) / 2 + 0.5,
                        inst.y1
                    )
                    strikeout_annot = page.add_redact_annot(strikeout_rect, fill=(1, 0, 0))
                    page.apply_redactions()
                    page.insert_text((strikeout_rect.x0-2, strikeout_rect.y0-3), original_text, fontsize=10, color=(0, 0, 0), rotate=270)
                    page.insert_text((strikeout_rect.x0 + 10, strikeout_rect.y0), replacement_text, fontsize=12, color=(1, 0, 0), rotate=270)
                    
        document.save(output_path)
        document.close()

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

    # To replace the Clean_copy with Redline_copy
    def find_and_replace(pdf_path: str, output_path: str, replacements: List[Tuple[str, str]]):

        document = pymupdf.open(pdf_path)

        for page_num in range(len(document)):
            page = document.load_page(page_num)
            for original_text, replacement_text in replacements:

                text_instances = page.search_for(original_text) # search for the clean_copy in PDF
                text_info = extract_text_with_font_info(pdf_path, original_text) # Extract the font information
                font , size = find_word_font_info(text_info, original_text) # Extracting the font style and font size
                styled_text = fontstyle.apply(replacement_text, font) # Applying the font style and font size to the Redline_copy
                split_text = styled_text[0 : len(styled_text) - 3] # to resolve the "[0m" error

                for inst in text_instances:
                    page.insert_text((inst.x0 - 8, inst.y0), split_text, fontsize=size,color = (1,0,0),rotate = 270)
        
        document.save(output_path)
        document.close()

    # Extracting the Info from Excel
    def process_excel_and_replace(excel_file: pd.DataFrame, folder_path: str, output_folder: str):
        df = excel_file
        
        for index, row in df.iterrows():
            part_number = row['Part_Number']
            banner_copy = row['Clean_copy']
            redline_copy = row['Redline_copy']
            
            pdf_path = os.path.join(folder_path, f"{part_number}.pdf")
            intermediate_pdf = os.path.join(output_folder, f"{part_number}_intermediate.pdf")
            output_pdf = os.path.join(output_folder, f"{part_number}_Redline.pdf")

            rev_replace(pdf_path, intermediate_pdf, [(banner_copy, redline_copy)])
            find_and_replace(intermediate_pdf, output_pdf, [(banner_copy, redline_copy)])
            os.remove(intermediate_pdf)

    # Streamlit interface
    col = st.columns((1, 6, 1))
    with col[0]:
        st.write("")
    with col[1]:
        st.title("Redline Automation Tool")
    with col[2]:
        st.write("")

    st.write("Please provide the necessary inputs to process your PDF files.")

    folder_path = st.text_input("Enter the folder path where the PDF files are located:")
    excel_file = st.file_uploader("Upload the Excel file:", type=["xlsx"])
    output_folder = st.text_input("Enter the folder path where the output PDF files should be saved:")

    if st.button("Start Process"):
        if folder_path and excel_file and output_folder:
            excel_data = pd.read_excel(excel_file)
            process_excel_and_replace(excel_data, folder_path, output_folder)
            st.success("Process completed successfully!")
        else:
            st.error("Please provide all inputs.")

    if st.button("Go Back"):
        switch_page('main')