import json
import os
import shutil
from lib.config_manager import create_directories, load_config
from lib.pdf_processing import convert_pdfs_to_images
from lib.presentation import add_image_to_presentation, is_duplicate


config = load_config()
pdf_input_dir = config['pdf_input_dir']
output_dir = config['output_dir']


def process_directory_files():

    if os.path.isdir(pdf_input_dir):
        # Iterate over all subdirectories of input_dir
        for root, dirs, files in os.walk(pdf_input_dir):
            for dir in dirs:
                dir_path = os.path.join(root, dir)

                # Get the sender name from the directory name
                sender = dir

                sender_dir, attachments_dir, text_dir, attachments_img_dir, text_img_dir = create_directories(
                    output_dir, sender)

            attachment_filepaths = []

            # Iterate over all PDF files in the directory
            for file in os.listdir(dir_path):
                if file.endswith(".pdf"):
                    # Save the filepath
                    attachment_filepaths.append(os.path.join(dir_path, file))

            attachment_images = convert_pdfs_to_images(
                attachment_filepaths, attachments_img_dir)

            for image in attachment_images:
                if not is_duplicate(image, sender):
                    add_image_to_presentation(
                        image, sender, "Per Post")

            return
    else:
        print(
            f'Skipping processing of .pdf files. {pdf_input_dir} is not a directory')
