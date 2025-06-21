# Nitori Catalog Automation Tool

![Automation](https://img.shields.io/badge/Type-Automation-blue) 
![Python](https://img.shields.io/badge/Python-3.8%2B-green) 
![License](https://img.shields.io/badge/License-MIT-orange)

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
- Required Python packages (listed in `requirements.txt`)

## Installation

### 1. Clone this repository

```bash
git clone https://github.com/chimboy19/-Nitori-Catalog-Automation-System
cd -Nitori-Catalog-Automation-System
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Install Tesseract OCR

#### Windows
Download installer from [Tesseract GitHub](https://github.com/tesseract-ocr/tesseract)

#### Mac (using Homebrew)
```bash
brew install tesseract
```

### 4. Set your Gemini API key

Create a `.env` file in the root directory with the following content:

```
GEMINI_API_KEY=your_api_key_here
```

## Usage

1. Place all input PDF files into the `input_pdfs/` folder  
2. Place your master product list Excel file in the root directory  
3. Run the automation script:

```bash
python main.py
```

## Customization

To adapt for different catalog layouts or company-specific formats:

- Edit the `OCR_PROMPT` to reflect the structure you expect in each catalog entry  
- Change the `REQUIRED_MASTER_COLS` list to match your Excel's actual column headers  
- Modify the `match_products()` function to tweak fuzzy matching logic

## Troubleshooting

- Review logs in `debug/nitori_automation.log` for detailed error tracing  
- To enable verbose logging, set the logging level to `DEBUG` in the code  
- If Tesseract fallback OCR fails, confirm the installation path and verify it’s accessible via command line  
- For best results, use high-resolution and well-scanned PDF files

## Limitations

- Output quality depends heavily on the input PDF clarity and layout consistency  
- Unusual or inconsistent catalog designs may require manual configuration  
- Pages with excessive density or complex design may reduce OCR accuracy  
- Does not support visual PDF annotation with bounding boxes (e.g., drawing product areas)

## License

This project is licensed under the MIT License.  
See the [LICENSE](LICENSE) file for details.
