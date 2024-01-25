import json
import os
import shutil
from lib.config_manager import create_directories, load_config
from lib.pdf_processing import convert_pdfs_to_images
from lib.presentation import add_image_to_presentation, is_duplicate


config = load_config()
scanned_input_dir = config['scanned_input_dir']
output_dir = config['output_dir']


def process_directory_files():
    if scanned_input_dir:
        # Iterate over all subdirectories of input_dir
        for root, dirs, files in os.walk(scanned_input_dir):
            for dir in dirs:
                dir_path = os.path.join(root, dir)

                # Read the JSON file
                with open(os.path.join(dir_path, "info.json"), 'r') as f:
                    info = json.load(f)

                sender = info['From']
                sender_dir, attachments_dir, text_dir, attachments_img_dir, text_img_dir = create_directories(
                    output_dir, sender)

            # copy the json file to the sender directory
            shutil.copy(os.path.join(dir_path, "info.json"), sender_dir)

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
                        image, sender)

            return
    else:
        print('No scanned_input_dir configured in config.yaml')
        exit()
