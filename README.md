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

## Input

### Email .eml files

Put .eml files in the directory specified in the configuration.

### Scanned documents as PDF

Place the scanned PDF documents in a subdirectory within the directory specified for scanned documents in the configuration. Group these documents by the submitter. The subdirectory's name is flexible and serves only for grouping purposes. In this subdirectory, create an info.json file with the appropriate content.

```json
{
  "From": "Name of the sender",
  "Date": "2000-01-01"
}
```

This is used to add the document source information to the .pptx file.

## Running the script

```bash
python eml2pptx.py
```

## Output

The output .pptx files are placed in the directory specified in the configuration.

## Adding more input files (.eml and .pdf)

The script will only process files that have not been processed before. To process more files, simply add them to the input directories and rerun the script. New slides will be added to the existing .pptx file. Unfortunately, there is no way to insert a new slide at an arbitrary position in a .pptx file, so the slides will be appended to the end of the file.
