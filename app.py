import streamlit as st
import os
import sys
import warnings

from PIL import Image
from PIL import Image, ImageOps, ImageDraw

# from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from fontTools import ttLib

import fitz
import numpy as np
import pandas as pd
import zipfile
import traceback
import re
from fuzzywuzzy import fuzz

import base64
import streamlit.components.v1 as components

if not sys.warnoptions:
    warnings.simplefilter("ignore")

st.set_page_config(layout="wide", page_title="PDF generation")
#col1, col2 = st.columns((0.3, 0.7))
#col2.write("# Physician Bio Generator")

cwd = os.getcwd()

col1, col_spacer_1, col2, col_spacer_2, = st.columns((1, 0.2, 0.6, 0.4))

# try:
#     with open("link_to_folder.txt", "r") as f:
#         link = f.read()
#     link = link.strip()
# except:
#     link = ''

# if link == '':
#     link = col1.text_input("Please enter the link to the folder with images:")

# text_size = 10 #8.25
# section_text_size = 9.5
# header_fontsize = 21.3

#W
font_size_1 = col2.slider("Name font size:", min_value=10.0, max_value=15.0, value=11.5, step = 0.1)
font_size_2 = col2.slider("Credentials font size:", min_value=6.0, max_value=12.6, value=7.0, step = 0.1)
font_size_4 = col2.slider("Title font size:", min_value=6.0, max_value=12.6, value=8.0, step = 0.1)
font_size_3 = col2.slider("Contact info font size:", min_value=6.0, max_value=11.0, value=7.7, step = 0.1)
# image_size = col2.slider("Image size:", min_value=100, max_value=160, value=120, step = 1)
col2.write('###')

# ColorMinMax = st.markdown(''' <style> div.stSlider > div[data-baseweb = "slider"] > div[data-testid="stTickBar"] > div {
#     background: rgb(1 1 1 / 0%); } </style>''', unsafe_allow_html = True)


# Slider_Cursor = st.markdown(''' <style> div.stSlider > div[data-baseweb="slider"] > div > div > div[role="slider"]{
#     background-color: rgb(14, 38, 74); box-shadow: rgb(14 38 74 / 20%) 0px 0px 0px 0.2rem;} </style>''', unsafe_allow_html = True)

    
# Slider_Number = st.markdown(''' <style> div.stSlider > div[data-baseweb="slider"] > div > div > div > div
#                                 { color: rgb(14, 38, 74); } </style>''', unsafe_allow_html = True)
    

# col = f''' <style> div.stSlider > div[data-baseweb = "slider"] > div > div {{
#     background: linear-gradient(to right, rgb(1, 183, 158) 0%, 
#                                 rgb(1, 183, 158) {text_size}%, 
#                                 rgba(151, 166, 195, 0.25) {text_size}%, 
#                                 rgba(151, 166, 195, 0.25) 100%); }} </style>'''

# ColorSlider = st.markdown(col, unsafe_allow_html = True)   

#col1.write("#### Please upload the Excel file and images:")
uploaded_files = col1.file_uploader(
    label="Please upload the Excel file:",
    #label_visibility = "collapsed",
    type=[
        "xlsx",
        "png",
        "jpg",
        "jpeg",
    ],
    accept_multiple_files=True,
)

file_dict = {}
for uploaded_file in uploaded_files:
    file_dict[uploaded_file.name] = uploaded_file


#item_dict = dict(zip(index_list, item_list))

#@st.cache_data(ttl= 300, show_spinner = False, max_entries = 500)
# def get_file_content(photo_loc):
#     content = item_df[item_df['name']==photo_loc].iloc[0]['item'].content()
#     return content

# if link == '':
#     item_df = pd.DataFrame()
# else:
#     item_df = get_item_df(link)
#     # try:
#     #     item_df = get_item_df(link)
#     # except:
#     #     print('Error when processing the link')
#     #     st.error(
#     #                 "\n\n\nError when processing the link"
#     #             )


def download_doc_button(object_to_download, download_filename):
    """
    Generates a link to download the given object_to_download.
    Params:
    ------
    object_to_download:  The object to be downloaded.
    download_filename (str): filename and extension of file. e.g. mydata.csv,
    Returns:
    -------
    (str): the anchor tag to download object_to_download
    """
    try:
        # some strings <-> bytes conversions necessary here
        b64 = base64.b64encode(object_to_download.read()).decode()

    except AttributeError as e:
        b64 = base64.b64encode(object_to_download).decode()

    dl_link = f"""
    <html>
    <head>
    <title>Start Auto Download file</title>
    <script src="http://code.jquery.com/jquery-3.2.1.min.js"></script>
    <script>
    $('<a href="data:application/pdf;base64,{b64}" download="{download_filename}">')[0].click()
    </script>
    </head>
    </html>
    """
    return dl_link


# col3.image('./logo.png')


# z = zipfile.ZipFile("inputs.zip")
# z.extractall('.')

# excel_files = [
#     filename for filename in os.listdir(".") if filename.endswith("xlsx")
# ]
# input_file_name = excel_files[0]

col1.write("#")
#col3, col4 = st.columns((0.75, 1.45))
col3, col4 = st.columns((2.5, 1))
col3.write("#")
# col3.write("#")
# col3.write("#")

excel_file_name = ''

#missing_image_filenames = False
excel_output = io.BytesIO()
excel_dataframes = {}

#@st.cache(suppress_st_warning=True, ttl= 300, show_spinner = False, max_entries = 10)
#@st.cache_data(ttl= 180, show_spinner = False, max_entries = 10)
def create_pdfs(file_dict, text_size, section_text_size, header_fontsize):
    excel_file_uploaded = False
    number_of_excel_files = 0
    for key in file_dict.keys():
        if key.endswith("xlsx"):
            excel_file_name = key
            excel_file_uploaded = True
            number_of_excel_files = number_of_excel_files + 1 
            #break

    if not excel_file_uploaded:
        #col1.error("Please upload the Excel file")
        st.stop()

    if number_of_excel_files > 1:
        col1.error("More than one Excel file uploaded")
        st.stop()

    xl = pd.ExcelFile(file_dict[excel_file_name])
    sheet_names = xl.sheet_names
    missing_image_filenames = False

    vertical_distance_same_field = 14.5
    vertical_distance_between_people = 48
    vertical_distance_between_fields_1 = 17
    vertical_distance_between_fields_2 = 20
    vertical_distance_between_fields_3 = 24

    #W
    font_path_1 = 'assets/Brandon Grotesque Bold.otf'
    font_path_2 = 'assets/BrandonGrotesque-Regular.otf'
    font_path_3 = 'assets/FilsonProBlack.otf'
    # font_path_4 = 'assets/MavenPro-Regular.ttf'
    # font_path_5 = 'assets/Maven Pro Light - 300.otf'


    font_1 = fitz.Font("font1", font_path_1) #fitz.Font("NoeDisplayBoldCheck", "NoeDisplay-Bold.ttf")
    font_2 = fitz.Font("font2", font_path_2) #fitz.Font("PoppinsRegularCheck", "Poppins-Light.otf")
    font_3 = fitz.Font("font3", font_path_3) #fitz.Font("NoeDisplayBoldCheck", "NoeDisplay-Bold.ttf")
    # font_4 = fitz.Font("font4", font_path_4) #f
    # font_5 = fitz.Font("font5", font_path_5)

    # font_1_char = ttLib.TTFont(font_path_1)
    # font_2_char = ttLib.TTFont(font_path_2)

    # # Load the TTF file
    # font = ttLib.TTFont('Avenir Black.ttf')

    # # Define the horizontal scaling factor
    # scaling_factor = 0.9  # Adjust this as needed

    # # Modify the glyf table to apply horizontal scaling
    # for glyph in font['glyf']:
    #     glyph.recalcBounds(font)
    #     glyph.scale((scaling_factor, 1))

    # # Save the modified font to a new file
    # font.save('Avenir Black adj.ttf')


    # text_x = 272.5
    # field_x = 82 #173.5

    horizontal_limit_threshold = 250 #570
    header_horizontal_limit_threshold = 608.5

    # def insert_text(text, x, y):
    #     text = text.replace("‐", "-").replace("\u200B", "")
    #     check_text_characters(text, font_2_char)
    #     rc = page.insert_text(
    #         fitz.Point(x, y),
    #         text,
    #         fontname= 'font2', #"PoppinsRegular",
    #         fontsize=text_size,
    #         color=(0.3254902, 0.33333333, 0.35686275),
    #     )

    def insert_text_(text, x, y, font_path = font_path_1, font_name = 'font1', font_size = font_size_1, color = (0.0627451, 0.18039216, 0.2627451), ):
        text = text.replace("‐", "-").replace("\u200B", "")
        check_text_characters(text, ttLib.TTFont(font_path))
        rc = page.insert_text(
            fitz.Point(x, y),
            text,
            fontname= font_name, #"PoppinsRegular",
            fontsize=font_size,
            color=color,
            #rotate = 180,
        )

    # def insert_name(text, x, y):
    #     text = text.replace("‐", "-").replace("\u200B", "")
    #     check_text_characters(text, font_1_char)
    #     rc = page.insert_text(
    #         fitz.Point(x, y),
    #         text,
    #         fontname= 'font1', 
    #         fontsize=16.5,
    #         color=(0.0627451, 0.18039216, 0.2627451),
    #     )

    class UnsupportedCharacterError(Exception):
        pass

    def check_text_characters(text, font):
        for char in text:
            if ord(char) not in font["cmap"].tables[0].cmap:
                print(
                    "\n\n\n{} sheet: Unsupported character '{}' (ascii number {}) in text '{}'".format(
                        sheet_name, char, ord(char), text
                    )
                )
                st.error(
                    "\n\n\n{} sheet: Unsupported character '{}' (ascii number {}) in text '{}'".format(
                        sheet_name, char, ord(char), text
                    )
                )
                raise UnsupportedCharacterError("Unsupported character")


    # def insert_field_name(text, x, y):
    #     rc = page.insert_text(
    #         fitz.Point(x, y),
    #         text,
    #         # fontname = 'UYEVBE+Montserrat-Regular',
    #         fontsize=section_text_size,
    #         color=(0.0627451, 0.18039216, 0.2627451),
    #     )

    def split_substring_into_appropriate_length(
        my_string, my_font, my_font_size, input_x=82, text_type="text"
    ):
        lines = []

        if text_type == "header":
            threshold = header_horizontal_limit_threshold
        else:
            threshold = horizontal_limit_threshold

        max_text_length = threshold - input_x
        number_of_words = len(my_string.split(" "))

        starting_word_index = 0

        for i in range(1, number_of_words + 1):
            current_line = " ".join(my_string.split(" ")[starting_word_index:i])
            if my_font.text_length(current_line, my_font_size) < max_text_length:
                i = i + 1
            else:
                lines.append(
                    " ".join(my_string.split(" ")[starting_word_index : i - 1])
                )
                starting_word_index = i - 1

        lines.append(" ".join(my_string.split(" ")[starting_word_index:i]))
        return lines

    def split_string_into_appropriate_length(
        my_string, my_font, my_font_size, input_x=82, text_type="text"
    ):
        # my_string = my_string.replace('\n', '')
        # substring_list = my_string.split(';')
        my_string = my_string.replace(";", "\n")
        substring_list = my_string.split("\n")
        # substring_list = re.split('; |\n', my_string)

        substring_list = [substring for substring in substring_list if substring != ""]
        substring_list = [substring.strip() for substring in substring_list]
        substring_list = [
            split_substring_into_appropriate_length(
                substring, my_font, my_font_size, input_x, text_type
            )
            for substring in substring_list
        ]
        lines = sum(substring_list, [])
        return lines


    class MissingImageError(Exception):
        pass

    # def create_framed_photo(
    #     photo_loc, left_offset_ratio=0, upper_offset_ratio=0, side_length_ratio=0
    # ):
    #     filler_size = 470
    #     background_size = 576
    #     placement = int(np.round((background_size - filler_size) / 2))

    #     #im = Image.open(photo_loc)
    #     if photo_loc in file_dict:
    #         im = Image.open(file_dict[photo_loc]).convert("RGBA")
    #     elif (item_df.empty == False) & (photo_loc in item_df['name'].values):
    #         #item_index = item_df[item_df['name']==photo_loc].iloc[0]['index']
    #         try:
    #             content = get_file_content(photo_loc)
    #         except:
    #             st.error("Can't connect to Box. To use the app please upload the photos manually.")
    #             st.stop()
    #         #content = item_dict[item_index].content()
    #         # for item in items:
    #         #     if item.name == photo_loc:
    #         #         content = item.content()
            
    #         #content = item_df[item_df['name']==photo_loc].iloc[0]['item'].content()
    #         im = Image.open(io.BytesIO(content)).convert("RGBA")
    #     else:
    #         print(
    #             "\n\n\n{} sheet: Missing image file '{}'".format(
    #                 sheet_name, photo_loc
    #             )
    #         )
    #         raise MissingImageError("Missing image file")
    #     width = im.size[0]
    #     height = im.size[1]

    #     # if side_length_ratio > 1:
    #     #   side_length_ratio = 1
    #     # if upper_offset_ratio > 1:
    #     #   upper_offset_ratio = 1
    #     # if left_offset_ratio > 1:
    #     #   left_offset_ratio = 1

    #     side = min(width, height)
    #     # if side_length_ratio == 0:
    #     #   side = min(width,height)
    #     # else:
    #     #   side = min(width,height)
    #     #   side = side * side_length_ratio

    #     left = 0  # width * left_offset_ratio
    #     upper = 0  # height * upper_offset_ratio

    #     # im = im.crop((left, upper, left+side, upper+side))

    #     # im = im.resize((filler_size, filler_size))
    #     # bigsize = (im.size[0] * 3, im.size[1] * 3)
    #     # mask = Image.new('L', bigsize, 0)
    #     # draw = ImageDraw.Draw(mask)
    #     # draw.ellipse((0, 0) + bigsize, fill=255)
    #     # mask = mask.resize(im.size, Image.ANTIALIAS)
    #     # im.putalpha(mask)

    #     # output = ImageOps.fit(im, mask.size, centering=(0.5, 0.5))
    #     # output.putalpha(mask)
    #     # #output.save('output.png')

    #     background = Image.open("fw.png")#.convert('RGBA')
    #     background = background.resize((background_size, background_size))
    #     background.paste(im, (placement, placement), im)

    #     background = background.resize((300, 300))
    #     #background.save("fr.png")
        
    #     #save pillow image to bytes
    #     imgByteArr = io.BytesIO()
    #     background.save(imgByteArr, format='PNG')
    #     #imgByteArr = imgByteArr.getvalue()
    #     return imgByteArr


    def insert_image(
        image_loc, upper, left_offset_ratio=0, upper_offset_ratio=0, side_length_ratio=0
    ):
        side = image_size #150 #120
        left = 42
        rect = fitz.Rect(left, upper, left + side, upper + side)
        byte_image = create_framed_photo(
            image_loc,
            left_offset_ratio=left_offset_ratio,
            upper_offset_ratio=upper_offset_ratio,
            side_length_ratio=side_length_ratio,
        )
        page.insert_image(rect, 
                        #filename="fr.png",
                        stream = byte_image)
        # img_xref = img_xref + 1

    def get_split_string_or_empty_list(input_string, my_font = font_1, my_font_size = 15.5, input_x = 12, ):
        if pd.isna(input_string) == False:
            split_list = split_string_into_appropriate_length(
                input_string, my_font, my_font_size, input_x
            )
        else:
            split_list = []
        return split_list

    # def insert_person(
    #     y,
    #     # image_loc,
    #     name_string,
    #     info_string_1,
    #     info_string_2,
    #     # school_string,
    #     # residency_string,
    #     # fellowship_string,
    #     # specialty_string,
    #     image_left_offset_ratio=0,
    #     image_upper_offset_ratio=0,
    #     image_side_length_ratio=0,
    #     check_y_needed=False,
    # ):
    #     # delimiter = ';'
    #     # school_list = school_string.split(delimiter)
    #     # school_list = [school.strip() for school in school_list]

    #     # residency_list = residency_string.split(delimiter)
    #     # residency_list = [residency.strip() for residency in residency_list]

    #     # fellowship_list = fellowship_string.split(delimiter)
    #     # fellowship_list = [fellowship.strip() for fellowship in fellowship_list]

    #     # specialty_list = specialty_string.split(delimiter)
    #     # specialty_list = [specialty.strip() for specialty in specialty_list]

    #     # education_list = split_string_into_appropriate_length(education_string, font_2, 8.25)
    #     # experience_list = split_string_into_appropriate_length(experience_string, font_2, 8.25)
    #     # school_list = split_string_into_appropriate_length(school_string, font_2, 8.25)
    #     # residency_list = split_string_into_appropriate_length(residency_string, font_2, 8.25)
    #     # fellowship_list = split_string_into_appropriate_length(fellowship_string, font_2, 8.25)
    #     # specialty_list = split_string_into_appropriate_length(specialty_string, font_2, 8.25)

    #     info_list_1 = get_split_string_or_empty_list(info_string_1)
    #     info_list_2 = get_split_string_or_empty_list(info_string_2)
    #     # school_list = get_split_string_or_empty_list(school_string)
    #     # residency_list = get_split_string_or_empty_list(residency_string)
    #     # fellowship_list = get_split_string_or_empty_list(fellowship_string)
    #     # specialty_list = get_split_string_or_empty_list(specialty_string)

    #     # y = y + vertical_distance_between_people
    #     if not check_y_needed:
    #         # insert_image(
    #         #     image_loc,
    #         #     y - 17,
    #         #     left_offset_ratio=image_left_offset_ratio,
    #         #     upper_offset_ratio=image_upper_offset_ratio,
    #         #     side_length_ratio=image_side_length_ratio,
    #         # )  # y-30)
    #         #insert_name(name_string, field_x, y)
    #         insert_text_(name_string, field_x, y, font_path_1, 'font1', 16.5, color = (0.0627451, 0.18039216, 0.2627451), )

    #     for i, info_1 in enumerate(info_list_1):
    #         if i == 0:
    #             y = y + vertical_distance_between_fields_1
    #             # if not check_y_needed:
    #             #     insert_field_name("EDUCATION:", field_x, y)
    #         else:
    #             y = y + vertical_distance_same_field
    #         if not check_y_needed:
    #             #insert_text(info_1, field_x, y)
    #             insert_text_(info_1, field_x, y, font_path_2, 'font2', 8, color = (0.0627451, 0.18039216, 0.2627451), )

    #     for i, info_2 in enumerate(info_list_2):
    #         if i == 0:
    #             y = y + vertical_distance_between_fields_2
    #             # if not check_y_needed:
    #             #     insert_field_name("EXPERIENCE:", field_x, y)
    #         else:
    #             y = y + vertical_distance_same_field
    #         if not check_y_needed:
    #             #insert_text(info_2, field_x, y)
    #             insert_text_(info_2, field_x, y, font_path_2, 'font2', 8, color = (0.0627451, 0.18039216, 0.2627451), )


    #     return y

    # def insert_header(my_page, header_string, header_fontsize=21.3):
    #     # delimiter = ';'
    #     # header_list = header_string.split(delimiter)
    #     # header_list = [header_part.strip() for header_part in header_list]
    #     header_string = header_string.replace("‐", "-").replace("\u200B", "")
    #     check_text_characters(header_string, font_1_char)
    #     header_list = split_string_into_appropriate_length(
    #         header_string, font_1, header_fontsize, 206, "header"
    #     )
    #     header_y = 43
    #     for header_part in header_list:
    #         rc = page.insert_text(
    #             fitz.Point(206, header_y),
    #             header_part,
    #             fontname="NoeDisplayBold",
    #             fontsize=header_fontsize,
    #             color=(1, 1, 1),
    #         )
    #         vertical_distance_within_header = 25
    #         header_y = header_y + vertical_distance_within_header

    # uploaded_image_filenames = [filename for filename in file_dict.keys() if not filename.endswith('.xlsx')]
    # if item_df.empty == False:
    #     additional_filenames = [filename for filename in list(item_df['name'].unique()) if not filename.endswith('.xlsx')]
    # else:
    #     additional_filenames = []
    # image_filenames= list(set(uploaded_image_filenames + additional_filenames))

    # def get_closest_image_filename(row):
    #     fullname = row['Name']
        
    #     max_score = 0
    #     closest_match = ''
    #     for image_filename in image_filenames:
    #         filename = image_filename
    #         if '.png' in filename:
    #             filename = filename.replace('_Rev.png', '')
    #             filename = filename.replace('_rev.png', '')
    #         else:
    #             filename = filename.rsplit('.', 1)[0]

    #         score = fuzz.token_sort_ratio(filename, fullname)
    #         if score > max_score:
    #             max_score = score
    #             closest_match = image_filename
    #     if max_score >= 60:
    #         return closest_match
    #     else:
    #         return ''

    page_dict = {}
    pdf_dict = {}
    successfully_created_list = []

    person_number = 0
    total_number_of_people = 0
    for sheet_name in sheet_names:
        try:
            data_df = pd.read_excel(file_dict[excel_file_name], sheet_name=sheet_name)
            total_number_of_people += len(data_df)
        except Exception:
            continue

    for sheet_name in sheet_names:
        try:
            doc = fitz.open("assets/temp.pdf")
            page = doc[1]
            page.wrap_contents()
            # page.insert_font("NoeDisplayBold", "NoeDisplay-Bold.ttf")
            # page.insert_font("PoppinsRegular", "Poppins-Light.otf")
            # page.insert_font('PoppinsRegular2', 'Poppins-Medium-1.otf')
            # page.insert_font('PoppinsRegular3', 'Poppins-Light.otf')
            # page.insert_font('PoppinsRegular4', 'Poppins-Thin.otf')
            # page.insert_font('NoeDisplayBold2', 'NoeDisplay-Bold.ttf')

            # data_df = pd.read_excel('input.xlsx', sheet_name = sheet_name)

            data_df = pd.read_excel(file_dict[excel_file_name], sheet_name=sheet_name)
            data_df = data_df.dropna(how="all")

            sections = [
                "Education",
                "Experience",
                "Medical School",
                "Residency",
                "Fellowship",
                "Specialty",
                "Image filename",
                'Image left offset ratio',
                'Image upper offset ratio',
                'Image side length ratio',
                'Image side length ratio',
                'Header font size',
            ]
            for section in sections:
                if section not in data_df.columns:
                    data_df[section] = np.nan

            excel_dataframes[sheet_name] = data_df                                                      

            start_y = 98 #152.5
            test_offset = 8

            y = start_y - vertical_distance_between_people
            person_on_page_index = 0
            previous_person_position = -1000

            page_number = 0

            progress_text = "Creating PDFs. Please wait..."
            my_bar = st.progress(0, text=progress_text)

            for i, r in data_df.iterrows():
                #print('i')
                try:
                    my_bar.progress(person_number/total_number_of_people, text=progress_text)
                    person_number += 1

                    name = r["Name"]
                    #name = name.upper()
                    # title = r["Title"]
                    # name = title.upper()

                    doc = fitz.open("assets/temp.pdf")
                    #doc = fitz.open("sample.pdf")


                    ###############################
                    #PAGE 1

                    #W
                    page = doc[0]
                    page.wrap_contents()
                    page.insert_font("font1", font_path_1)
                    page.insert_font("font2", font_path_2)
                    page.insert_font("font3", font_path_3)
                    # page.insert_font("font4", font_path_4)
                    # page.insert_font("font5", font_path_5)

                    x = 103
                    color = (1, 1, 1) # (0.0627451, 0.18039216, 0.2627451) #
                    # color = (0, 0, 0)

                    insert_text_(name, x, 23, font_path_1, 'font3', font_size_1, color = color, ) #  (0.0627451, 0.18039216, 0.2627451)
                    insert_text_(r["Credentials"], x+1, 30.5, font_path_2, 'font2', font_size_2, color = color, )
                    insert_text_(r["Title"], x, 44, font_path_3, 'font1', font_size_4, color = color, )

                    x = 119
                    insert_text_(r["Info field 1 (phone)"], x, 84.5, font_path_2, 'font2', font_size_3, color = color, )
                    insert_text_(r["Info field 2 (email)"], x, 100.5, font_path_2, 'font2', font_size_3, color = color, )
                    insert_text_(r["Info field 3 (url)"], x, 116.5, font_path_2, 'font2', font_size_3, color = color, )

                    # for i, info in enumerate(get_split_string_or_empty_list(r["Title"], my_font = font_3, my_font_size = font_size_2, input_x = 82, )):
                    #     y = 61
                    #     if i > 0:
                    #         y = y + 8.5
                    #     insert_text_(info, 82, y, font_path_1, 'font3', font_size_2, color = (1.0, 1.0, 1.0), )

                    # for i, info_1 in enumerate(info_list_1):
                    #     if i == 0:
                    #         y = y + vertical_distance_between_fields_1
                    #     else:
                    #         y = y + vertical_distance_same_field
                    #     insert_text_(info_1, 93, 93, font_path_2, 'font2', font_size_3, color = (1.0, 1.0, 1.0), )

                    # for i, info_2 in enumerate(info_list_2):
                    #     if i == 0:
                    #         y = y + vertical_distance_between_fields_2
                    #     else:
                    #         y = y + vertical_distance_same_field
                    #     insert_text_(info_2, 93, 110, font_path_2, 'font2', font_size_3, color = (1.0, 1.0, 1.0), )

                    ###############################
                    #PAGE 3

                    # page = doc[3]
                    # page.insert_font("font1", font_path_1)  
                    # page.insert_font("font2", font_path_2)
                    # page.insert_font("font3", font_path_3)
                    # y = start_y - vertical_distance_between_people

                    # info_list_1 = get_split_string_or_empty_list(r["Email"], my_font = font_2, my_font_size = font_size_3, input_x = 12, )
                    # info_list_2 = get_split_string_or_empty_list(r["Phone"], my_font = font_2, my_font_size = font_size_3, input_x = 12, )
                    # #info_list_3 = get_split_string_or_empty_list(r["Title"], 12)

                    # insert_text_(name, 12, 49.5, font_path_1, 'font1', font_size_1, color = (0.050980392156862744, 0.7058823529411765, 0.8156862745098039), ) #  (0.0627451, 0.18039216, 0.2627451)

                    # #insert_text_(r["Title"], 12, 59, font_path_1, 'font3', 8.6, color = (0.050980392156862744, 0.7058823529411765, 0.8156862745098039), )
                    # for i, info in enumerate(get_split_string_or_empty_list(r["Title"], my_font = font_3, my_font_size = font_size_2, input_x = 12, )):
                    #     y = 61
                    #     if i > 0:
                    #         y = y + 8.5
                    #     insert_text_(info, 12, y, font_path_1, 'font3', font_size_2, color = (0.050980392156862744, 0.7058823529411765, 0.8156862745098039), )

                    # for i, info_1 in enumerate(info_list_1):
                    #     if i == 0:
                    #         y = y + vertical_distance_between_fields_1
                    #     else:
                    #         y = y + vertical_distance_same_field
                    #     insert_text_(info_1, 24, 112, font_path_2, 'font2', font_size_3, color = (0.9450980392156862, 0.3137254901960784, 0.3058823529411765), ) # 

                    # for i, info_2 in enumerate(info_list_2):
                    #     if i == 0:
                    #         y = y + vertical_distance_between_fields_2
                    #     else:
                    #         y = y + vertical_distance_same_field
                    #     insert_text_(info_2, 24, 129, font_path_2, 'font2', font_size_3, color = (0.9450980392156862, 0.3137254901960784, 0.3058823529411765), )

                    my_bar.empty()

                    # name = "{} {}".format(
                    #     r["First Name"],
                    #     r["Last Name"],
                    # )
                    # if pd.isnull(r["Image left offset ratio"]):
                    #     left_offset = 0
                    # else:
                    #     left_offset = r["Image left offset ratio"]
                    # if pd.isnull(r["Image upper offset ratio"]):
                    #     upper_offset = 0
                    # else:
                    #     upper_offset = r["Image upper offset ratio"]
                    # if pd.isnull(r["Image side length ratio"]):
                    #     side_length = 0
                    # else:
                    #     side_length = r["Image side length ratio"]

                    # minimum_space_between_people = 150
                    # if y - previous_person_position < minimum_space_between_people:
                    #     y = previous_person_position + minimum_space_between_people
                    # previous_person_position = y

                    # check_y = insert_person(
                    #     y,
                    #     #image_loc=r["Image filename"],
                    #     name_string=name,
                    #     info_string_1=r["Email"],
                    #     info_string_2=r["Phone"],
                    #     # school_string=r["Medical School"],
                    #     # residency_string=r["Residency"],
                    #     # fellowship_string=r["Fellowship"],
                    #     # specialty_string=r["Specialty"],
                    #     # check_y_needed = True,
                    #     # image_left_offset_ratio=left_offset,
                    #     # image_upper_offset_ratio=upper_offset,
                    #     # image_side_length_ratio=side_length,
                    #     check_y_needed=True,
                    # )


                    # y = insert_person(
                    #     y,
                    #     #image_loc=r["Image filename"],
                    #     name_string=name,
                    #     info_string_1=r["Email"],
                    #     info_string_2=r["Phone"],
                    #     # school_string=r["Medical School"],
                    #     # residency_string=r["Residency"],
                    #     # fellowship_string=r["Fellowship"],
                    #     # specialty_string=r["Specialty"],
                    #     # check_y_needed = True,
                    #     # image_left_offset_ratio=left_offset,
                    #     # image_upper_offset_ratio=upper_offset,
                    #     # image_side_length_ratio=side_length,
                    #     # check_y_needed = True,
                    # )
                    # #person_on_page_index = person_on_page_index + 1

                    # insert_text('qwertyuiopasdfghjklzxcvbnm — & test', 170, y+20)
                    # page_dict["page_{}".format(page_number)] = doc.tobytes()
                    # #doc.save("page_{}.pdf".format(page_number))

                    # #open fitz file from bytes
                    

                    # #doc = fitz.open("page_0.pdf")
                    # doc = fitz.open(stream=page_dict["page_0"])
                    # for i in range(1, page_number + 1):
                    #     #doc2 = fitz.open("page_{}.pdf".format(i))
                    #     doc2 = fitz.open(stream=page_dict["page_{}".format(i)])
                    #     doc.insert_pdf(doc2)
                    #     doc2.close()

                    file_name = "{}.pdf".format(sheet_name)
                    metadata_dict = doc.metadata
                    header_text = name
                    metadata_dict["title"] = header_text.replace(
                        ";", ""
                    )  #'New Specialists Joining'
                    doc.set_metadata(metadata_dict)

                    pdf_dict[name] = doc.tobytes(
                        garbage=4,
                        deflate=True,
                        clean=True,
                        deflate_images=True,
                        deflate_fonts=True,
                    )
                    print("\nSuccessfully created {}.pdf".format(name))
                    successfully_created_list.append(file_name)
                    #st.success("Successfully created {}".format(file_name))
                except UnsupportedCharacterError:
                    print("PDF couldn't be generated for the sheet {}, unsupported characters for person: {}".format(sheet_name, name))
                    #st.error("The person's name: " + name)
                    col1.error("PDF couldn't be generated for the sheet {}, unsupported characters for person: {}".format(sheet_name, name))
                    continue
                except MissingImageError:
                    print("PDF couldn't be generated for the sheet {}, image not uploaded: {}".format(sheet_name, r["Image filename"]))
                    col1.error("PDF couldn't be generated for the sheet {}, image not uploaded: {}".format(sheet_name, r["Image filename"]))
                    continue
                except Exception:
                    print("PDF couldn't be generated for the sheet {}, person: {}".format(sheet_name, name))
                    #st.error("The person's name: " + name)
                    col1.error("PDF couldn't be generated for the sheet {}, person: {}".format(sheet_name, name)  + traceback.format_exc())
                    continue

        except Exception:
            print("PDF couldn't be generated for the sheet '{}'. Error when processing data in the sheet:".format(sheet_name))
            #st.write('#')
            col1.error("PDF couldn't be generated for the sheet '{}'. Error when processing data in the sheet: \n\n\n\n".format(sheet_name) + traceback.format_exc())
            #st.error(traceback.format_exc())
            traceback.print_exc()
            continue

    if missing_image_filenames:
        for sheet_name in sheet_names:
            data_df = pd.read_excel(file_dict[excel_file_name], sheet_name=sheet_name)
            data_df = data_df.dropna(how="all")

            sections = [
                "Image filename",
            ]
            for section in sections:
                if section not in data_df.columns:
                    data_df[section] = np.nan

            data_df['Image filename'] = data_df.apply(lambda x: x['Image filename'] if pd.isna(x['Image filename']) == False else get_closest_image_filename(x), 
                                                                                        axis = 1)
            excel_dataframes[sheet_name] = data_df

        with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
            for sheet_name in sheet_names:
                excel_dataframes[sheet_name].to_excel(
                    writer, engine="xlsxwriter", index=False, sheet_name=sheet_name,
                )


    # buf = io.BytesIO()
    # with open("Avera St. Luke's.pdf", "rb") as f:
    #     buf.write(f.read())
    #     buf.seek(0)

    # components.html(
    #     download_doc_button(buf, "report.pdf"),
    #     height=0,
    # )
    # buf.close()

    success_string = ""
    for file_name in successfully_created_list:
        success_string = success_string +  "Successfully created " + file_name + "\n\n"  

    return pdf_dict, missing_image_filenames, excel_file_name


with col2:
    # with st.spinner("Creating PDFs..."):
        # # delete a list of files
        # temporary_exceptions = ["Avera St. Luke's.pdf", "Westover Hills.pdf"]
        # files_to_remove = [
        #     filename for filename in os.listdir(".") if (filename.endswith("pdf") and filename != "temp.pdf" and filename not in temporary_exceptions)
        # ]
        # print(files_to_remove)
        # for f in files_to_remove:
        #     try:
        #         os.remove(f)
        #     except OSError:
        #         print(f)
        #         pass

    pdf_dict, missing_image_filenames, excel_file_name = create_pdfs(file_dict, font_size_1, font_size_2, font_size_3)


# st.download_button(
#     label="Download PDFs",
#     data=pdf_dict[list(pdf_dict.keys())[0]],
#     file_name='test.pdf',
#     mime='application/pdf',
# )


if missing_image_filenames == True:
    col4.download_button(
        label="Download Excel",
        data=excel_output,
        file_name='{} (with filled in filenames).xlsx'.format(excel_file_name.rsplit('.',1)[0]),
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
elif len(pdf_dict) == 1:
    sheet_name = list(pdf_dict.keys())[0]
    col4.download_button(
        label="Download PDF",
        data=pdf_dict[sheet_name],
        file_name='{}.pdf'.format(sheet_name),
        mime='application/pdf',
    )
elif len(pdf_dict) > 1:
    z = io.BytesIO()
    with zipfile.ZipFile(z, "w") as zf:
        for sheet_name in pdf_dict:
            zf.writestr("{}.pdf".format(sheet_name), pdf_dict[sheet_name])

    # with open("inputs.zip", "rb") as fp:
    col4.download_button(
        label="Download PDFs",
        data=z,
        file_name='Created PDFs.zip',
        mime='application/zip',
    )


# if col4.button("Create PDFs"):

#             #col1.success(success_string)

#     for sheet_name in pdf_dict:
#         #st.success(sheet_name)
#         file_name = "{}.pdf".format(sheet_name).replace("'", "").replace("\"", "").replace(":", "").replace(";", "").replace("&", "")
#         try:
#             components.html(
#                 download_doc_button(pdf_dict[sheet_name], file_name), #sheet_name + ".pdf"),
#                 height=0,
#             )
#         except Exception as e:
#             st.error(e)

#             # components.html(
#             #     download_doc_button(pdf_dict["Avera St. Luke's"], "{}.pdf".format("Avera St. Luke's")), #sheet_name + ".pdf"),
#             #     height=0,
#             # )


#             # buf = io.BytesIO()
#             # with open("Avera St. Luke's.pdf", "rb") as f:
#             #     buf.write(f.read())
#             #     buf.seek(0)

#             # buf = io.BytesIO()
#             # buf.write(doc.tobytes())

#             # components.html( 
#             #     download_doc_button(pdf_dict[sheet_name], "report2.pdf"),
#             #     height=0,
#             # )
#             #buf.close()