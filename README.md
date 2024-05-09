![image](https://github.com/csb21jb/TextHarvester/assets/94072917/90e08671-0733-4a71-b2ba-207aac5f5f68)


# TextHarvester

![Python](https://img.shields.io/badge/Python-3.6%2B-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

## Overview

`TextHarvestor` is a powerful Python script that automates the extraction of text from various document formats, including PowerPoint, PDF, Word, and plain text files. The extracted content is then compiled into a single, organized Word document for easy reference and analysis.

### Features

- Extracts text from **PowerPoint (.pptx)**, **PDF (.pdf)**, **Word (.docx, .doc)**, and **Text (.txt)** files.
- Compiles all extracted text into a single organized Word document.
- Includes robust error handling and logging mechanisms.
- Progress bar for tracking the extraction process.
- Automated dependency installation to simplify setup.

## Installation

### Prerequisites
- Python 3.6+
- Pip (Python package manager)

### Setup Instructions
1. Clone the repository:
    ```bash
    git clone https://github.com/csb21jb/TextHarvester.git
    cd TextHarvester
    ```

2. Install dependencies:
    ```bash
    pip3 install python-pptx pdfplumber python-docx tqdm colorama
    ```
    ```bash
    pip3 install extract-msg
    ```
    ```bash
    pip3 install textract --no-deps
    ```
## Usage

1. Place the Python script (e.g., `TextHarvester.py`) in a directory containing the documents you want to extract text from.

2. Run the Python script:
    ```bash
    python3 TextHarvester.py
    ```

3. The extracted text will be saved in a file named `combined_output.docx`.

## Example Output

Here's an example of the expected output structure in the `combined_output.docx` file:


https://github.com/csb21jb/TextHarvester/assets/94072917/0803bbfa-6cba-4e71-8090-c05eb95b7d11


