# Purpose: Extracts email attachments from .eml files and saves them to a directory
from lib.directory_files_processing import process_directory_files
from lib.email_processing import process_eml_files
from lib.info import print_duplicate_senders, print_new_pages_start
from lib.presentation import count_slides, get_all_senders
from lib.xlsx_processing import update_excel_file

# conda install conda-forge::weasyprint

from lib.presentation import add_headers, create_presentation_from_dict, save_presentation

starting_page_count = count_slides()

process_eml_files()
process_directory_files()
prs = create_presentation_from_dict()
add_headers(prs)
save_presentation(prs)

update_excel_file(get_all_senders())

print_new_pages_start(count_slides(), starting_page_count)

print_duplicate_senders(prs)
