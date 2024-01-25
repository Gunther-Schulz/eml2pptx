# Purpose: Extracts email attachments from .eml files and saves them to a directory
from lib.directory_files_processing import process_directory_files
from lib.email_processing import process_eml_files
from lib.info import print_duplicate_senders, print_new_pages_start

# conda install conda-forge::weasyprint

from lib.presentation import add_headers, create_presentation_from_dict, save_presentation

starting_page_count = 0


process_eml_files()
process_directory_files()
prs = create_presentation_from_dict()
page_count = add_headers(prs)
save_presentation(prs)

print_new_pages_start(page_count, starting_page_count)

print_duplicate_senders(prs)
