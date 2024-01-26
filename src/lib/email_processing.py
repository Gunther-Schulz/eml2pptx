from email import policy
from email.parser import BytesParser
from email.utils import parseaddr
import os
import shutil
from lib.config_manager import get_email_date, is_in_blacklist, load_config, create_directories, sanitize_filename
from mailparser_reply import EmailReplyParser
from lib.pdf_processing import convert_pdfs_to_images, create_pdf
from lib.presentation import add_image_to_presentation, is_duplicate


config = load_config()
output_dir = config['output_dir']
eml_input_dir = config['eml_input_dir']

# The ordfer ist important. The first content type will be used if multiple are found
text_content_types = ['text/plain', 'text/html']


def process_eml_files():
    if os.path.isdir(eml_input_dir):
        for filename in os.listdir(eml_input_dir):
            if filename.endswith('.eml'):
                process_single_eml_file(filename, output_dir)
        return
    else:
        print(
            f'Skipping processing of .eml files. {eml_input_dir} is not a directory')


def process_single_eml_file(filename, output_dir):
    with open(f'./eml/{filename}', 'rb') as f:
        msg = BytesParser(policy=policy.default).parse(f)

    sender = parseaddr(msg['From'])[1]
    sender_dir, attachments_dir, text_dir, attachments_img_dir, text_img_dir = create_directories(
        output_dir, sender)

    # copy the eml file to the sender directory
    shutil.copy(f'./eml/{filename}', sender_dir)

    message_content = get_html_content(msg)

    if message_content is None:
        raise Exception(
            f'No HTML/Text content found in {filename} for sender {sender}')

    filename = sanitize_filename(filename)

    text_filepath = create_pdf(message_content, text_dir, filename)
    text_images = convert_pdfs_to_images([text_filepath], text_img_dir)

    attachment_filepaths = save_attachments(msg, attachments_dir)

    attachment_images = convert_pdfs_to_images(
        attachment_filepaths, attachments_img_dir)

    all_images = text_images + attachment_images

    for image in all_images:
        if not is_duplicate(image, sender):
            add_image_to_presentation(
                image, sender)
    return


def get_html_content(msg):
    for type in text_content_types:
        if msg.is_multipart():
            for part in msg.iter_parts():
                if part.get_content_type() == type:
                    # if Content-Disposition is attachment, skip
                    if part.get('Content-Disposition') and part.get('Content-Disposition') != "inline":
                        continue
                    reply = remove_quoted(part.get_content())
                    if reply:
                        return {"type": type, "content": reply}
                    return {"type": type, "content": part.get_content()}
                elif part.is_multipart():
                    return get_html_content(part)
        else:
            if msg.get_content_type() == type:
                if msg.get('Content-Disposition') and msg.get('Content-Disposition') != "inline":
                    continue
                reply = remove_quoted(msg.get_content())
                if reply:
                    return {"type": type, "content": reply}
                return {"type": type, "content": msg.get_content()}

    return None


def remove_quoted(email_message):
    email_message = EmailReplyParser(
        languages=['de', 'en']).parse_reply(text=email_message)
    return email_message


def save_attachments(msg, output_dir):
    filepaths = []
    if msg.is_multipart():
        for part in msg.iter_parts():
            if part.get_content_type() == 'application/pdf':
                fp = handle_pdf(part, msg, output_dir)
                if fp is not None:
                    filepaths.append(fp)
            elif part.get_content_type() == 'application/zip':
                handle_zip(part, msg, output_dir)
            elif part.get('Content-Disposition') and part.get_filename() is not None:
                fp = handle_other_attachments(part, msg, output_dir)
                if fp is not None:
                    filepaths.append(fp)
    return filepaths


def handle_pdf(part, msg, output_dir):
    payload = part.get_payload(decode=True)
    if payload is not None:
        attachment_filename = part.get_filename()
        if is_in_blacklist(attachment_filename):
            return None
        formatted_date = get_email_date(msg)
        filename, file_extension = os.path.splitext(attachment_filename)
        filename_with_date = f"{filename}_{formatted_date}{file_extension}"
        filepath = os.path.join(output_dir, filename_with_date)
        with open(filepath, 'wb') as f:
            f.write(payload)
        return filepath


def handle_zip(part, msg, output_dir):
    # TODO: handle zip files. unpack them, save the files and add them to the presentation
    payload = part.get_payload(decode=True)
    if payload is not None:
        attachment_filename = part.get_filename()
        if is_in_blacklist(attachment_filename):
            return None
        filepath = os.path.join(output_dir, attachment_filename)
        print(
            f'Found zip file. Saving to: {attachment_filename} in path {filepath}')
        with open(filepath, 'wb') as f:
            f.write(payload)


def handle_other_attachments(part, msg, output_dir):
    payload = part.get_payload(decode=True)
    if payload is not None:
        attachment_filename = part.get_filename()
        if is_in_blacklist(attachment_filename):
            return None
        filepath = os.path.join(output_dir, attachment_filename)
        if attachment_filename.endswith('.pdf'):
            print(
                f'Found PDF attachment without application/pdf MIME type. Saving to: {attachment_filename} in path {filepath}')
            return handle_pdf(part, msg, output_dir)
        print(
            f'Found non-PDF attachment. Saving to: {attachment_filename} in path {filepath}')
        with open(filepath, 'wb') as f:
            f.write(payload)
