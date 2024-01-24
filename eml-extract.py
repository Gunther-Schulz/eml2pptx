import os
import email
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
import shutil
from weasyprint import HTML
from pdf2image import convert_from_path
from PIL import Image
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptx.util import Mm, Cm
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.shapes import PP_PLACEHOLDER
import re
import hashlib
import json

# TODO: implenet adding slides from scanned files (no email)

# conda install conda-forge::weasyprint 

# check version of pptx library
# pip show python-pptx

presentation_filename = 'presentation.pptx'

text_types = ['text/plain', 'text/html']
eml_input_dir = './eml'
scanned_input_dir = './scanned'
output_dir = './output'
font_path = "/Library/Fonts/Arial.ttf"
page_count = 0
slides_dict = {}
regex_right_header = re.compile(r"Seite \d+/\d+") # regex to match the right header

# Create output directory if it doesn't exist
if not os.path.exists(output_dir):
    os.makedirs(output_dir)


# Check if the PowerPoint file already exists
if os.path.exists(presentation_filename):
    # If it does, open it
    prs = Presentation(presentation_filename)

else:
    prs = Presentation()

# Set slide width and height to A4 portrait dimensions (210mm x 297mm)
prs.slide_width = Mm(297)
prs.slide_height = Mm(210)

def image_basename(image):
    return os.path.splitext(os.path.basename(image))[0]

def hash_image_name(image_name):
    name = image_basename(image_name)
    hasher = hashlib.md5()
    hasher.update(name.encode('utf-8'))
    unique_representation = hasher.hexdigest()
    # return "hash_image_" + unique_representation
    return image_name

def hash_sender_name(sender):  
    hasher = hashlib.md5()
    hasher.update(sender.encode('utf-8'))
    unique_representation = hasher.hexdigest()
    # return "hash_sender_" + unique_representation
    return sender

# def write_default_json_to_notes(slide):
#     # prs = Presentation()
#     # title_slide_layout = prs.slide_layouts[0]
#     # blank_slide_layout = prs.slide_layouts[6]

#     # title_slide = prs.slides.add_slide(title_slide_layout)
#     # title = title_slide.shapes.title
#     # title.text = "Title"

#     # notes_slide = title_slide.notes_slide #The only new line of code

#     # blank_slide = prs.slides.add_slide(blank_slide_layout)
#     # notes_slide = blank_slide.notes_slide
#     # notes_slide.notes_text_frame.text = "foo"

#     notes_slide = slide.notes_slide

#     text_frame = notes_slide.notes_text_frame

#     notes_placeholder = notes_slide.notes_placeholder

 
#     text_frame.text = 'foobar'

#     default_for_json = {"hash_image_name": "default", "hashed_sender_name": "default"}
#     json_hashed_image_name = json.dumps(default_for_json)
#     slide.notes_slide.notes_text_frame.text = json_hashed_image_name

def write_default_json_config_to_text_frame(slide):
    # create a new text box, not a note
    text_frame = slide.shapes.add_textbox(Cm(0.1), Cm(0), Cm(0.1), Cm(0.1)).text_frame # Parameters are left, top, width, height
    # write default json to it
    default_for_json = {"hash_image_name": "default", "hashed_sender_name": "default"}
    json_hashed_image_name = json.dumps(default_for_json)
    text_frame.text = json_hashed_image_name
    text_frame.paragraphs[0].font.size = Pt(1)
    text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
 
def get_json_text_box(slide):
    # search all text boxes and find the first one that contains json
    regex_json = re.compile(r"^{.*}$")
    for shape in slide.shapes:
        if shape.has_text_frame:
            text_frame = shape.text_frame
            if regex_json.match(text_frame.text):
                return text_frame
    return None

def write_config_to_text_frame(slide, key, value):
    json_text_box = get_json_text_box(slide)
    if json_text_box is None:
        return
    # get the json from the text box and update it
    existing_json = json.loads(json_text_box.text)
    existing_json[key] = value
    # write the updated json back to the text box
    json_text_box.text = json.dumps(existing_json)
    json_text_box.paragraphs[0].font.size = Pt(1)
    json_text_box.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

def read_config_from_text_frame(slide, key):
    json_text_box = get_json_text_box(slide)
    if json_text_box is None:
        return None
    # get the json from the text box
    existing_json = json.loads(json_text_box.text)
    return existing_json.get(key, None)

# def write_to_slide_note(slide, key, value):
#     # Read the existing content

#     # temp_prs = Presentation()
#     # blank_slide_layout = temp_prs.slide_layouts[6]
#     # notes_slide = slide.notes_slide #The only new line of code

#     # blank_slide = temp_prs.slides.add_slide(blank_slide_layout)
#     # notes_slide = blank_slide.notes_slide
#     # notes_slide.notes_text_frame.text = "foo"

#     existing_content = slide.notes_slide.notes_text_frame.text
#     # existing_content = None
#     existing_json = json.loads(existing_content) if existing_content else {}

#     # Update the JSON
#     existing_json[key] = value

#     # Write it back
#     slide.notes_slide.notes_text_frame.text = json.dumps(existing_json)

# def read_from_slide_note(slide, key):
#     # print(slide.has_notes_slide)
#     # Read the existing content
#     existing_content = slide.notes_slide.notes_text_frame.text
#     existing_json = json.loads(existing_content) if existing_content else {}

#     # Update the JSON
#     return existing_json[key]

def sanitize_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', '_', filename)


def get_html_content(msg):
    if msg.is_multipart():
        for part in msg.iter_parts():
            if part.get_content_type() in text_types:
                return part.get_content()
            elif part.is_multipart():
                return get_html_content(part)
    else:
        if msg.get_content_type() == 'text/html':
            return msg.get_content()

    return None

from PIL import Image
from PIL import ImageChops

def add_header_left(slide):
    bold = False
    size = 11
    # Add a header to the slide
    header = slide.shapes.add_textbox(Cm(1), 0, prs.slide_width - Cm(2), Cm(1))
    tf = header.text_frame
    tf.text = "Abwägungsvorschlag Träger öffentlicher Belange, B-Plan XX “XXXX“"
    tf.paragraphs[0].font.size = Pt(size)
    tf.paragraphs[0].font.bold = bold
    tf.paragraphs[0].alignment = PP_ALIGN.LEFT

def add_header_right(slide, page_number, total_pages=page_count):
    bold = False
    size = 10
    header = slide.shapes.add_textbox(Cm(1), 0, prs.slide_width - Cm(2), Cm(1))
    tf = header.text_frame
    tf.text = f'Seite {page_number}/{total_pages}'
    tf.paragraphs[0].font.size = Pt(size)
    tf.paragraphs[0].font.bold = bold
    tf.paragraphs[0].alignment = PP_ALIGN.RIGHT

def crop_whitespace(image):
    # Crop the image to remove whitespace
    bg = Image.new(image.mode, image.size, image.getpixel((0,0)))
    diff = ImageChops.difference(image, bg)
    diff = ImageChops.add(diff, diff, 2.0, -100)
    bbox = diff.getbbox()
    if bbox:
        return image.crop(bbox)

def add_image(slide, image_filepath):
    # Open the image and get its size
    image = Image.open(image_filepath)
    image_width, image_height = image.size
    aspect_ratio = image_width / image_height

    # Define the maximum height and width
    max_height = Cm(18)
    max_width = Cm(13)

    # Calculate the height and width, maintaining the aspect ratio
    if aspect_ratio > 1:
        # Image is wider than it is tall
        width = max_width
        height = width / aspect_ratio
    else:
        # Image is taller than it is wide
        height = max_height
        width = height * aspect_ratio

    # Add the image to the slide
    top = Cm(1.6)
    left = Cm(1)
    pic = slide.shapes.add_picture(image_filepath, left, top, height=height, width=width)

    # Crop the image
    pic.crop_left = 0
    pic.crop_top = 0
    pic.crop_right = 0
    pic.crop_bottom = 0

# Function that adds a string to the top left of the slide
def add_text(slide, text):
    left = top = width = height = Cm(1)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.text = text
    tf.paragraphs[0].font.size = Pt(6)

def add_divider_line(slide, prs):
    left = prs.slide_width // 2
    top = Cm(0.7)
    end_top = prs.slide_height
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, left, top, left, end_top)
    line.line.color.rgb = RGBColor(0, 0, 0)
    line.line.width = Pt(1)
    line.shadow.inherit = False
    line.shadow.visible = False

    # Add horizontal divider line 1 cm from the top
    left = Cm(0.5)
    top = Cm(0.7)
    end_left = prs.slide_width - Cm(0.5)
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, left, top, end_left, top)
    line.line.color.rgb = RGBColor(0, 0, 0)
    line.line.width = Pt(1)
    line.shadow.inherit = False
    line.shadow.visible = False


def add_text_box(slide):
    left = Cm(16)
    top = Cm(1)
    width = Cm(11)
    height = Cm(20)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    p = tf.add_paragraph()
    p.text = "Zur Kenntnis genommen."
    p.font.bold = False
    p.font.size = Pt(10)
    p.alignment = PP_ALIGN.LEFT

def remove_title_placeholder(slide):
    for shape in slide.placeholders:
        if shape.is_placeholder:
            phf = shape.placeholder_format
            if phf.idx == 0 and phf.type == PP_PLACEHOLDER.TITLE:
                sp = shape._element
                sp.getparent().remove(sp)

def create_directories(output_dir, sender):
    # Create a new directory with the same name as the sender's email address
    sender_dir = f'{output_dir}/{sender}'
    os.makedirs(sender_dir, exist_ok=True)
    
    # Create a subdirectory called "attachments" to store attachments
    attachments_dir = f'{sender_dir}/attachments'
    os.makedirs(attachments_dir, exist_ok=True)
    
    # Create a subdirectory called "text" to store HTML files
    text_dir = f'{sender_dir}/text'
    os.makedirs(text_dir, exist_ok=True)
    
    # Create a subdirectory called "img" in subdirectory "attachments" to store images
    attachments_img_dir = f'{attachments_dir}/img'
    os.makedirs(attachments_img_dir, exist_ok=True)
    
    # Create a subdirectory called "img" in subdirectory "text" to store images
    text_img_dir = f'{text_dir}/img'
    os.makedirs(text_img_dir, exist_ok=True)

    return sender_dir, attachments_dir, text_dir, attachments_img_dir, text_img_dir
    
def get_email_date(msg):
    email_date = parsedate_to_datetime(msg['Date'])
    # Format the date and time as YYYYMMDD_HHMMSS
    formatted_date_time = email_date.strftime('%Y-%m-%d_%H_%M_%S')
    return formatted_date_time

def save_attachments(msg, output_dir):
    filepaths = []
    if msg.is_multipart():
        for part in msg.iter_parts():
            if part.get_content_type() not in text_types:
                payload = part.get_payload(decode=True)
                if payload is not None:
                    attachment_filename = part.get_filename()
                    formatted_date =  get_email_date(msg)
                    filename, file_extension = os.path.splitext(attachment_filename)
                    # Include the date in the filename
                    filename_with_date = f"{filename}_{formatted_date}{file_extension}"
                    filepath = os.path.join(output_dir, filename_with_date)
                    with open(filepath, 'wb') as f:
                        f.write(payload)
                    filepaths.append(filepath)
    return filepaths

def convert_pdfs_to_images(pdf_filepaths, output_dir):
    images = []
    for pdf_filepath in pdf_filepaths:
        # Convert the PDF to images
        pdf_images = convert_from_path(pdf_filepath)
        # Save each image under the output directory
        for i, pdf_image in enumerate(pdf_images):
            pdf_image = crop_whitespace(pdf_image)
            image_filepath = f'{output_dir}/{os.path.basename(pdf_filepath)}_{i}.png'
            pdf_image.save(image_filepath, 'PNG')
            images.append(image_filepath)
    return images

def create_pdf(html_content, output_dir, filename):
    filepath = f'{output_dir}/{filename}.pdf'
    # Create a PDF of the email and save it under the new directory
    HTML(string=html_content).write_pdf(filepath)
    return filepath

def process_directory_files(input_dir, output_dir):
    # Iterate over all subdirectories of input_dir
    for root, dirs, files in os.walk(input_dir):
        for dir in dirs:
            dir_path = os.path.join(root, dir)
            
            # Read the JSON file
            with open(os.path.join(dir_path, "info.json"), 'r') as f:
                info = json.load(f)

            sender = info['From']
            sender_dir, attachments_dir, text_dir, attachments_img_dir, text_img_dir = create_directories(output_dir, sender)

            # copy the json file to the sender directory
            shutil.copy(os.path.join(dir_path, "info.json"), sender_dir)

            attachment_filepaths = []

            # Iterate over all PDF files in the directory
            for file in os.listdir(dir_path):
                if file.endswith(".pdf"):
                    # Save the filepath
                    attachment_filepaths.append(os.path.join(dir_path, file))

            attachment_images = convert_pdfs_to_images(attachment_filepaths, attachments_img_dir)

            for image in attachment_images:
                
                hashed_image_name = hash_image_name(image)
                hashed_sender_name = hash_sender_name(sender)

                if not is_duplicate_image(hashed_image_name):
                    add_image_to_presentation(image, hashed_image_name, hashed_sender_name)

            return


def process_eml_files(input_dir, output_dir):
    for filename in os.listdir(input_dir):
        if filename.endswith('.eml'):
            process_single_eml_file(filename, output_dir)
    return

def process_single_eml_file(filename, output_dir):
    with open(f'./eml/{filename}', 'rb') as f:
        msg = BytesParser(policy=policy.default).parse(f)
    
    sender = email.utils.parseaddr(msg['From'])[1]
    sender_dir, attachments_dir, text_dir, attachments_img_dir, text_img_dir = create_directories(output_dir, sender)

    # copy the eml file to the sender directory
    shutil.copy(f'./eml/{filename}', sender_dir)

    html_content = get_html_content(msg)
    if html_content is None:
        print(f'No HTML content found in {filename}')
    filename = sanitize_filename(filename)

    text_filepath = create_pdf(html_content, text_dir, filename)
    text_images = convert_pdfs_to_images([text_filepath], text_img_dir)
    
    attachment_filepaths = save_attachments(msg, attachments_dir)
                    
    attachment_images = convert_pdfs_to_images(attachment_filepaths, attachments_img_dir)

    all_images = text_images + attachment_images

    for image in all_images:
        
        hashed_image_name = hash_image_name(image)
        hashed_sender_name = hash_sender_name(sender)

        if not is_duplicate_image(hashed_image_name):
            add_image_to_presentation(image, hashed_image_name, hashed_sender_name)

    return

def is_duplicate_image(hash_image_name):
    # return any(read_from_slide_note(slide, "hashed_image_name") == hashed_image_name for slide in prs.slides)
    return any(read_config_from_text_frame(slide, "hash_image_name") == hash_image_name for slide in prs.slides)

def add_image_to_presentation(image, hashed_image_name, hashed_sender_name):
    # Check if the sender already exists in the dictionary
    if hashed_sender_name in slides_dict:
        # Add the image to the sender's list
        slides_dict[hashed_sender_name].append(image)
    else:
        # If the sender does not exist, create a new list for them
        slides_dict[hashed_sender_name] = [image]


def create_presentation_from_dict(prs):
    i = 0
    for sender, images in slides_dict.items():
        for image in images:
            i += 1
            hashed_image_name = hash_image_name(image)
            # if not is_duplicate_image(hashed_image_name):
            # Add a new slide at the end
            print(f'Adding slide nr. {i} {image} to slide with hashes {hashed_image_name}')
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            # write_default_json_to_notes(slide)
            # write_to_slide_note(slide, "hashed_image_name", hashed_image_name)
            # write_to_slide_note(slide, "hashed_sender_name", sender)
            write_default_json_config_to_text_frame(slide)
            write_config_to_text_frame(slide, "hash_image_name", hashed_image_name)
            write_config_to_text_frame(slide, "hash_sender_name", sender)
            remove_title_placeholder(slide)
            add_header_left(slide)
            add_image(slide, image)
            add_text(slide, f'Quelle: {sender} - {image_basename(image)}')
            add_text_box(slide)
            add_divider_line(slide, prs)


def add_headers_and_print_hashes(prs):
    page_count = len(prs.slides)
    for i, slide in enumerate(prs.slides):
        # if right header exists, remove it
        for shape in slide.shapes:
            if shape.has_text_frame:
                text_frame = shape.text_frame
                if regex_right_header.match(text_frame.text):
                    slide.shapes._spTree.remove(shape._element)
        add_header_right(slide, i+1, page_count)

    # for i, slide in enumerate(prs.slides):
    #     # hashed_image_name = read_from_slide_note(slide,"hashed_image_name")
    #     # hashed_sender_name = read_from_slide_note(slide,"hashed_sender_name")
    #     hashed_image_name = read_config_from_text_frame(slide, "hash_image_name")
    #     hashed_sender_name = read_config_from_text_frame(slide, "hash_sender_name")
    #     print(f'{i+1}: {hashed_image_name} - {hashed_sender_name}')


def save_presentation(prs):
    prs.save(presentation_filename)

def group_consecutive_numbers(numbers):
    ranges = []
    for n in numbers:
        if not ranges or n > ranges[-1][-1] + 1:
            ranges.append([n])
        else:
            ranges[-1].append(n)
    return ['{}-{}'.format(r[0], r[-1]) if len(r) > 1 else str(r[0]) for r in ranges]


# Usage
process_eml_files(eml_input_dir, output_dir)
process_directory_files(scanned_input_dir, output_dir)
create_presentation_from_dict(prs)
add_headers_and_print_hashes(prs)
save_presentation(prs)

# go through all slides and add the hashed sender name to a list
all_senders = []
for i, slide in enumerate(prs.slides):
    hashed_image_name = read_config_from_text_frame(slide, "hash_image_name")
    hashed_sender_name = read_config_from_text_frame(slide, "hash_sender_name")
    all_senders.append(hashed_sender_name)
# update list so that if adjecent elments are the same, they are replaced by a single element
senders = [x for i, x in enumerate(all_senders) if i == 0 or x != all_senders[i-1]]
# show which elements appear more than once in th elist
duplicates = [item for item in senders if senders.count(item) > 1]
# remove duplicates from list
senders = list(dict.fromkeys(duplicates))

print("The following senders appear in more than one section in the presentation:")
for sender in senders:
    positions = [i+1 for i, x in enumerate(all_senders) if x == sender]
    grouped_positions = group_consecutive_numbers(positions)
    print(f'{sender} appears at pages {", ".join(grouped_positions)}')
