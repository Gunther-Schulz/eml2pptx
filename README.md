# eml2pptx

This utility converts .eml files into .pptx files, specifically tailored for TöB. It aims to simplify the process of creating presentations for TöB and facilitating commentary. The tool incorporates the response body from the .eml file into the .pptx file and appends any attachments as images.

# Setup

Create a conda environment from environment.yml

```bash
conda env create -f environment.yml
```

Activate the environment

```bash
conda activate eml2pptx
```

# Usage

## Configuration

The settings for this application are contained in a `config.json` file, which should be located in the same directory as the executable. If this file doesn't exist initially, running the script once will generate a default config.json. You can then modify this file as needed and rerun the script.

### Presentation file name

`presentation_file_name`

This is the name of the .pptx file that will be generated. The .pptx extension is automatically appended.

### Slide Headers

There are the `header_title` and `page_string` fields, which are used to set the title of the presentation and the page number string. The page number string is used to set the page number in the footer of the presentation. The page number string should contain three `%d` placeholders, which will be replaced with the current sender number, the current page number, and the total number of pages, respectively.

### Default comment

`default_comment`

This is the default comment that will be added to the slide. This is a placeholder for the user to add their own comments.

### PDF Blacklist

`pdf_blacklist`

This is an array of regexes that will be used to filter out PDFs that match any of the regexes. This is useful for filtering out PDFs that are not relevant to the presentation, such as if an email was replied to containing the mail and PDF attachments from the previous email.

### Input directories

`email_input_dir`

This is the directory where the .eml files are located.

`pdf_input_dir`

This is the directory where the PDF files are located. Usually scanned documents.

### Output directory

`output_dir`

This is the directory where all extracted attachments, the text body of the reply and the original .eml file will be placed. Also, image files that are generated from the PDFs will be placed here. They are grouped by submitter. The email address of the submitter is used as the name of the subdirectory in case of .eml files and the name of the subdirectory is used as the name of the submitter in case of PDF files.

### Color code slides by sender

`color_code_sender`

This is a boolean value that determines whether slides should be color-coded by the submitter (sender). If this is set to true, each sender will have a different color bar on the left of the slide. This is useful for quickly identifying when a new "Stellungnahem" starts.

## Input

### Email .eml files

Put .eml files in the directory specified in the configuration.

### Scanned documents as PDF

Place the scanned PDF documents in a subdirectory within the directory specified for scanned documents in the configuration. Group these documents by the submitter. The subdirectory's name is used as the submitter's name.

This is used to add the document source information to the .pptx file.

## Running the script

```bash
python eml2pptx.py
```

## Output

The output .pptx files are placed in the directory specified in the configuration.

## Updating the presentation

### Adding/removing input files (.eml and .pdf)

The script will only process files that have not been processed before. To process more files, simply add them to the input directories and rerun the script. New slides will be added to the existing .pptx file. Unfortunately, there is no way to insert a new slide at an arbitrary position in a .pptx file, so the slides will be appended to the end of the file. Also, the underlying python-pptx library handling the .pptx files does not support deleting slides, so if you want to remove slides, you will have to do so manually.

Additionally, if .eml files are removed, this will not be reflected in the presentation. Slides are never removed (not supported by python-pptx). If an .eml file is removed, it also needs to be removed from the .pptx file manually.
