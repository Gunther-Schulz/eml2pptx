
from html import escape
import os
from pdf2image import convert_from_path
from weasyprint import HTML
from PIL import Image, ImageChops
from lib.config_manager import load_config


config = load_config()


def convert_pdfs_to_images(pdf_filepaths, output_dir):
    images = []
    for pdf_filepath in pdf_filepaths:
        # Convert the PDF to images
        pdf_images = convert_from_path(pdf_filepath)
        # Save each image under the output directory
        for i, pdf_image in enumerate(pdf_images):
            pdf_image = crop_whitespace(pdf_image)
            image_filepath = f'{output_dir}/{os.path.basename(pdf_filepath)}_{i}.png'
            if pdf_image is not None:
                pdf_image.save(image_filepath, 'PNG')
                images.append(image_filepath)
    return images


def create_pdf(message_content, output_dir, filename):
    content = message_content["content"]
    type = message_content["type"]
    # Check if the content is HTML or plain text
    if type == "text/plain":
        # The content is plain text
        # Set a maximum width to fit the content within an A4 page
        content = escape(content).replace('\n', '<br>')
        content = '<p style="word-wrap: break-word; max-width: 595px; font-size: 14pt;">{}</p>'.format(
            content)
    filepath = f'{output_dir}/{filename}.pdf'
    # Create a PDF of the email and save it under the new directory
    HTML(string=content).write_pdf(filepath)
    return filepath


def crop_whitespace(image):
    # Crop the image to remove whitespace
    bg = Image.new(image.mode, image.size, image.getpixel((0, 0)))
    diff = ImageChops.difference(image, bg)
    diff = ImageChops.add(diff, diff, 2.0, -100)
    bbox = diff.getbbox()
    if bbox:
        # Modify the bounding box to keep the full width of the image
        bbox = (0, bbox[1], image.size[0], bbox[3])
        return image.crop(bbox)
