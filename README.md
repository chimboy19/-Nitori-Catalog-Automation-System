markdown

# Nitori Catalog Automation Tool

![Automation](https://img.shields.io/badge/Type-Automation-blue) ![Python](https://img.shields.io/badge/Python-3.8%2B-green) ![License](https://img.shields.io/badge/License-MIT-orange)

## Overview

This tool automates the processing of Nitori product catalogs by extracting product information from PDFs and matching it with master Excel data. It combines OCR (Optical Character Recognition) with AI-powered data extraction to streamline catalog processing workflows.

## Key Features

- **Multi-engine OCR Processing**: Uses Gemini AI for primary OCR with Tesseract as a fallback
- **Intelligent Data Matching**: Fuzzy matching algorithms to link extracted data with master records
- **PDF Annotation**: Automatically annotates source PDFs with colored markers based on confidence levels
- **Excel Integration**: Updates master Excel files with extracted data and layout references
- **Confidence-based Processing**: Handles matches differently based on confidence scores
- **Parallel Processing**: Utilizes multi-threading for efficient page processing
- **Comprehensive Reporting**: Generates JSON results, Excel reports, and annotated PDFs

## Prerequisites

- Python 3.8+
- Tesseract OCR installed (for fallback processing)
- Google Gemini API key
- Required Python packages (listed in requirements.txt)

## Installation

1. Clone this repository:
   ```bash
git clone [repository-url]

## Install dependencies:
## bash
pip install -r requirements.txt

## Install Tesseract OCR:

## Windows: Download installer from Tesseract GitHub

## Mac: Use Homebrew:bash

brew install tesseract
## Set your Gemini API key in .env file:
GEMINI_API_KEY=your_api_key_here

## Usage

Place PDF files in the input folder

Put the product list Excel file in the main directory

## Run the tool:bash
python main.py

## Customization

## To adapt for different catalog formats:

Modify the OCR_PROMPT to match your expected data structure

Update REQUIRED_MASTER_COLS with your Excel column names

Adjust matching logic in match_products() function

## Troubleshooting

Check debug/nitori_automation.log for detailed error information

Enable debug mode by setting logging level to DEBUG in code

Verify Tesseract installation path if fallback OCR fails

Ensure PDFs are clear scans with legible text

## Limitations

Performance depends on PDF quality and layout complexity

Non-standard catalog formats may require customization

Very dense pages may overwhelm OCR processing

Unable to handle PDF annotations (e.g., creating borders for product identification)

## License

This project is licensed under the MIT License. See LICENSE file for details.