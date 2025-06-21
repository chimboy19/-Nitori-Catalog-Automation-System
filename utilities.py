import os
import fitz  # PyMuPDF
import cv2
import numpy as np
from dotenv import load_dotenv
from rapidfuzz import fuzz, process
import google.generativeai as genai
import pandas as pd
import json
from typing import List, Dict, Optional, Tuple, Union
from concurrent.futures import ThreadPoolExecutor
import re
import time
import base64
from datetime import datetime
import logging
import pytesseract
import sys
from enum import Enum
from dataclasses import dataclass
import yaml

# Load environment variables
load_dotenv()

class AnnotationStyle(Enum):
    BOX = 1
    HIGHLIGHT = 2
    UNDERLINE = 3
    STRIKETHROUGH = 4

@dataclass
class AnnotationConfig:
    color: Tuple[float, float, float]
    thickness: float
    style: AnnotationStyle

# Config class with Gemini settings
class Config:
    def __init__(self):
        self.config_path = "config.yaml"
        self.load_defaults()
        
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    yaml_config = yaml.safe_load(f)
                    if yaml_config:
                        for key, value in yaml_config.items():
                            if key == "ANNOTATION_STYLES":
                                for style_name, style_dict in value.items():
                                    style_dict['style'] = AnnotationStyle[style_dict['style'].upper()]
                                    self.__dict__[key][style_name] = AnnotationConfig(**style_dict)
                            else:
                                self.__dict__[key] = value
            except Exception as e:
                logging.error(f"Error loading config.yaml: {e}. Using defaults.")
        
        # Set API key from environment
        if not self.GEMINI_API_KEY:
            self.GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
            if not self.GEMINI_API_KEY:
                logging.warning("GEMINI_API_KEY not found in config or environment")
        
        # Ensure directories exist
        os.makedirs(self.INPUT_DIR, exist_ok=True)
        os.makedirs(self.OUTPUT_DIR, exist_ok=True)
        os.makedirs(self.DEBUG_DIR, exist_ok=True)
    
    def load_defaults(self):
        # Gemini Settings
        self.GEMINI_API_KEY = None
        self.GEMINI_MODEL = "gemini-1.5-flash"
        self.GEMINI_SAFETY_SETTINGS = [
            {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"}
        ]
        self.GEMINI_GENERATION_CONFIG = {
            "temperature": 0.4,
            "top_p": 1,
            "top_k": 32,
            "max_output_tokens": 8192,
        }
        
        # Processing Parameters
        self.MAX_PAGES = None # Process all pages by default
        self.CHUNK_SIZE = 3 # Not explicitly used for GPT-4, mainly for batching smaller tasks
        self.MAX_WORKERS = 4 # For ThreadPoolExecutor
        self.REQUEST_DELAY = 1.0
        
        # Image Processing
        self.PDF_DPI = 300
        self.IMAGE_PREPROCESS = True
        
        # OCR Settings
        self.OCR_PROMPT = """Extract all product data as JSON with:
        - id (product code)
        - name (full product name) in japanese 
        - price (formatted as ¥12,345)
        - category (product type)
        - dimensions (WxDxH if available)
        - bbox [x1,y1,x2,y2] (PDF coordinates)"""
        
        self.REQUIRED_MASTER_COLS = [
            "ページ", "ｸﾞﾘｯﾄﾞNo", "ｻｲｽﾞ", "地域版", "対抗版", "掲載順", "差替理由", "差替先",
            "小分類CD", "担当バイヤー", "商品CD", "商品名", "商品名ｶﾅ", "掲載", "字組", "字組用商品名",
            "ｾｰﾙ売価", "売価", "大大分類 (家具／Hfa)", "売変開始日", "売変終了日", "材質名称", "色名称",
            "幅", "奥行", "高さ", "キャッチコピー", "枠付きコピー", "組立", "注記", "セール情報備考",
            "連絡事項", "展開店舗割合", "ファイル名１", "ファイル名２", "ファイル名３", "ファイル名４", "ファイル名５",
            "ピクト１", "ピクト２", "ピクト３", "ピクト４", "ピクト５", "内寸", "矢印寸法"
        ]
        
        # Rest of your existing defaults...
        self.INPUT_DIR = "input_pdfs"
        self.OUTPUT_DIR = "output"
        self.DEBUG_DIR = "debug"


        self.EXACT_MATCH_CONFIDENCE = 1.0
        self.GOOD_MATCH_CONFIDENCE = 0.8
        self.MIN_MATCH_CONFIDENCE = 0.5
        
        # Update Rules
        self.UPDATE_CORE_FIELDS = True
        self.OVERWRITE_EXISTING = False # Set to True to always overwrite fields if OCR has a value
        self.PRICE_TOLERANCE = 100 # Price difference in Yen to flag as mismatched
        
        # Enhanced Annotation Settings
        self.ANNOTATION_STYLES = {
            "high_confidence": AnnotationConfig(
                color=(0, 0.8, 0),  # Green (R,G,B float values 0-1)
                thickness=1.5,
                style=AnnotationStyle.BOX
            ),
            "medium_confidence": AnnotationConfig(
                color=(0.9, 0.6, 0),  # Orange
                thickness=1.0,
                style=AnnotationStyle.HIGHLIGHT
            ),
            "low_confidence": AnnotationConfig(
                color=(0.8, 0, 0),  # Red
                thickness=2.0,
                style=AnnotationStyle.BOX
            ),
            "needs_review": AnnotationConfig(
                color=(0.5, 0, 0.5),  # Purple
                thickness=1.5,
                style=AnnotationStyle.UNDERLINE
            )
        }
        
        # Excel Layout Reference Settings
        self.EXCEL_LAYOUT_REFERENCES = True
        self.EXCEL_COLORS = { # Hex colors for Excel
            "high_confidence": "00FF00",  # Green
            "medium_confidence": "FF9900",  # Orange
            "low_confidence": "FF0000",  # Red
            "needs_review": "9900FF"  # Purple
        }
        
        
        # ... (keep all other existing defaults)
        self.TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
# Initialize config and Gemini
config = Config()

# Set Tesseract path if it's specified in config
if sys.platform == "win32" and os.path.exists(config.TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = config.TESSERACT_PATH
elif sys.platform == "linux" or sys.platform == "darwin":
    # Assume Tesseract is in PATH for Linux/macOS or can be specified here
    pytesseract.pytesseract.tesseract_cmd = os.getenv("TESSERACT_PATH", "tesseract")
else:
    logging.warning("Tesseract path not configured. Tesseract OCR may not work.")


logging.basicConfig(
    filename=os.path.join(config.DEBUG_DIR, 'nitori_automation.log'),
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
genai.configure(api_key=config.GEMINI_API_KEY)

# ... (keep all other existing functions the same, except for the OCR functions shown above)

# ======================
# VALIDATION UTILITIES
# ======================
def validate_master_file(master_path: str) -> bool:
    """Validate the master Excel file has required columns"""
    try:
        master_df = pd.read_excel(master_path)
        missing_cols = [col for col in config.REQUIRED_MASTER_COLS 
                        if col not in master_df.columns]
        if missing_cols:
            logging.error(f"Master file '{master_path}' missing required columns: {missing_cols}")
            return False
        return True
    except Exception as e:
        logging.error(f"Master file validation failed for '{master_path}': {str(e)}")
        return False

# ======================
# IMAGE PROCESSING UTILITIES
# ======================
def preprocess_image(img: np.ndarray) -> np.ndarray:
    """Enhance image quality for better OCR results (generalized)"""
    try:
        # Convert to grayscale if it's a color image
        if len(img.shape) == 3 and img.shape[2] == 3:
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        else:
            gray = img # Already grayscale or single channel

        # Apply adaptive thresholding
        processed = cv2.adaptiveThreshold(
            gray, 255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY, 11, 2
        )
        
        # Apply a sharpening filter
        kernel = np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]])
        processed = cv2.filter2D(processed, -1, kernel)
        
        return processed
    except Exception as e:
        logging.warning(f"Image preprocessing failed: {str(e)}. Returning original image.")
        return img

def prepare_image_for_ocr(img: np.ndarray) -> str:
    """Convert image to base64 without data URL prefix for Gemini"""
    try:
        if len(img.shape) == 2:
            img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
        _, buffer = cv2.imencode('.jpg', img, [int(cv2.IMWRITE_JPEG_QUALITY), 85])
        return base64.b64encode(buffer).decode('utf-8')
    except Exception as e:
        logging.error(f"Image preparation failed: {str(e)}")
        raise

def clean_price(price_str: Union[str, float, int]) -> Optional[float]:
    """Extract numeric value from price string or convert existing numeric to float"""
    if price_str is None or pd.isna(price_str):
        return None
    
    if isinstance(price_str, (int, float)):
        return float(price_str)

    try:
        # Remove currency symbols (¥, $, €), commas, and spaces
        cleaned = re.sub(r'[^\d.]', '', str(price_str))
        if cleaned:
            return float(cleaned)
        return None
    except (ValueError, TypeError):
        return None

def normalize_bbox_to_pdf(bbox: List[Union[int, float]], page_height: float) -> List[float]:
    """Convert bbox coordinates from image (top-left origin) to PDF (bottom-left origin)"""
    try:
        if not isinstance(bbox, list) or len(bbox) != 4:
            logging.warning(f"Invalid bbox format received: {bbox}. Returning default.")
            return [0.0, 0.0, 100.0, 100.0]  # Return default bbox if invalid
            
        # Ensure all coordinates are floats
        x0, y0, x1, y1 = [float(coord) for coord in bbox]


        scale_factor = config.PDF_DPI / 72.0
        
        # Convert pixel coordinates from GPT to PDF points
        pdf_x0 = x0 / scale_factor
        pdf_y0 = y0 / scale_factor
        pdf_x1 = x1 / scale_factor
        pdf_y1 = y1 / scale_factor

        # Bboxes from GPT-4 Vision are typically top-left origin. PyMuPDF annotations also use top-left origin
        # for `fitz.Rect(x0, y0, x1, y1)` where y0 is top.
        # So, direct conversion after scaling should be enough.
        return [pdf_x0, pdf_y0, pdf_x1, pdf_y1]
    except Exception as e:
        logging.error(f"BBox normalization failed for bbox {bbox} and page_height {page_height}: {str(e)}")
        return [0.0, 0.0, 100.0, 100.0]  # Return default bbox on error
    
def pdf_to_images(pdf_path: str) -> List[Tuple[int, np.ndarray, float, float]]:
    """Convert PDF to images with original page dimensions and pixmap height for bbox normalization"""
    try:
        doc = fitz.open(pdf_path)
        images = []
        for i in range(doc.page_count):
            if config.MAX_PAGES and i >= config.MAX_PAGES:
                break
            page = doc.load_page(i)
            
            # Get pixmap at desired DPI
            pix = page.get_pixmap(dpi=config.PDF_DPI)
            
            # Convert pixmap to numpy array
            # PyMuPDF pix.samples is RGB. cv2.imdecode expects BGR usually, or convert to BGR for display.
            # For `prepare_image_for_ocr`, we convert to RGB for base64 encoding.
            # So, keep it RGB for now, or convert to BGR early for OpenCV processing
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
            
            # If the image is RGB (3 channels), convert to BGR for OpenCV processing
            if pix.n == 3:
                img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
            elif pix.n == 1: # Grayscale
                pass # No color conversion needed
            else:
                logging.warning(f"Unsupported number of channels ({pix.n}) for page {i+1}. Skipping image processing.")
                continue

            if config.IMAGE_PREPROCESS:
                img = preprocess_image(img)
            
            # Store page number, processed image, pixmap height (in pixels), and original page height (in points)
            images.append((i+1, img, pix.height, page.rect.height)) 
        doc.close()
        return images
    except Exception as e:
        logging.error(f"PDF to image conversion failed for '{pdf_path}': {str(e)}")
        raise
def gemini_ocr_extraction(page_num: int, img: np.ndarray, page_original_height_points: float) -> Optional[Dict]:
    """Primary OCR using Gemini Pro Vision"""
    try:
        img_base64 = prepare_image_for_ocr(img)
        img_data = base64.b64decode(img_base64)
        
        model = genai.GenerativeModel(config.GEMINI_MODEL)
        
        prompt_with_context = (
            config.OCR_PROMPT + 
            f"\nNote: The image is a rendering of PDF page {page_num}. " +
            "Bbox coordinates should correspond to the pixel dimensions of this image. " +
            "Respond ONLY with valid JSON in this exact format: {\"products\": [{\"id\": ..., \"name\": ..., etc}]}"
        )
        
        response = model.generate_content(
            contents=[
                prompt_with_context,
                {"mime_type": "image/jpeg", "data": img_data}
            ],
            generation_config=config.GEMINI_GENERATION_CONFIG,
            safety_settings=config.GEMINI_SAFETY_SETTINGS
        )
        
        if response.text:
            try:
                # Clean response text to extract JSON
                json_str = response.text.strip()
                if json_str.startswith("```json"):
                    json_str = json_str[7:]
                if json_str.endswith("```"):
                    json_str = json_str[:-3]
                return json.loads(json_str)
            except json.JSONDecodeError as e:
                logging.error(f"Failed to parse Gemini response as JSON: {response.text}")
                return None
        return None
        
    except Exception as e:
        logging.error(f"Gemini OCR failed on page {page_num}: {str(e)}")
        return None

def tesseract_ocr_fallback(page_num: int, img: np.ndarray, page_original_height_points: float) -> Optional[Dict]:
    """Fallback OCR using Tesseract"""
    try:
        # Preprocess specifically for Tesseract if not already done by general preprocess_image
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY) if len(img.shape) == 3 else img
        blurred = cv2.GaussianBlur(gray, (3, 3), 0)
        thresh = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
        
        # Configure Tesseract to output bounding box data and text
        # '--oem 3' (LSTM engine), '--psm 6' (Assume a single uniform block of text)
        # Use image_to_data for bbox information
        data = pytesseract.image_to_data(thresh, output_type=pytesseract.Output.DICT)
        text = pytesseract.image_to_string(thresh, config=r'--oem 3 --psm 6')

        products_found = []
        # Attempt to find bounding boxes for detected words and reconstruct potential product areas
        n_boxes = len(data['level'])
        for i in range(n_boxes):
            # Only consider words with high confidence
            if int(data['conf'][i]) > 70 and data['text'][i].strip():
                (x, y, w, h) = (data['left'][i], data['top'][i], data['width'][i], data['height'][i])
                
                # Simple heuristic to group words into potential product areas
                # This is a very basic fallback; GPT is much better for structured extraction
                product_text = data['text'][i].strip()
                
                # Attempt to extract product ID and price from this word/area
                product_id_match = re.search(r'[A-Z]{2}\d{4}', product_text)
                price_match = re.search(r'¥[\d,]+', product_text)

                if product_id_match or price_match:
                    
                    scale_factor_x = page_original_height_points / img.shape[0] # Roughly, assuming aspect ratio maintained
                    scale_factor_y = page_original_height_points / img.shape[0]
                   
                    bbox_pixels = [x, y, x + w, y + h]
                    bbox_pdf_points = normalize_bbox_to_pdf(bbox_pixels, img.shape[0]) # Use img.shape[0] as pix.height

                    products_found.append({
                        "id": product_id_match.group() if product_id_match else "",
                        "name": product_text, # Tesseract doesn't easily delineate names
                        "price": price_match.group() if price_match else "",
                        "bbox": bbox_pdf_points,
                        "page_num": page_num
                    })
        
        # If no specific products found, maybe return a single entry with overall text
        if not products_found and text.strip():
             # Create a dummy bbox for the whole page if no specific products could be extracted
             # This requires knowing the original PDF page dimensions
             # For a fallback, if a product is not found with a bbox, it's better to not add a dummy bbox
             # but rather indicate that the data is not well-structured.
             logging.info(f"Tesseract found text on page {page_num} but couldn't structure it.")
             return {
                "products": [{
                    "id": "",
                    "name": text.strip()[:100] + "..." if len(text.strip()) > 100 else text.strip(),
                    "price": "",
                    "bbox": [0,0,0,0], # Indicate no meaningful bbox
                    "page_num": page_num
                }]
             }
        
        if products_found:
            return {"products": products_found}
        return None
    except pytesseract.TesseractNotFoundError:
        logging.error("Tesseract is not installed or not found in PATH. Please install Tesseract OCR or set TESSERACT_PATH in config.yaml/environment variable.")
        return None
    except Exception as e:
        logging.error(f"Tesseract OCR failed on page {page_num}: {str(e)}")
        return None


# Update the enhanced_ocr_extraction to use gemini_ocr_extraction instead of gpt_ocr_extraction
def enhanced_ocr_extraction(page_num: int, img: np.ndarray, page_original_height_points: float) -> Optional[Dict]:
    """Try Gemini first, fallback to Tesseract if needed"""
    try:
        result = gemini_ocr_extraction(page_num, img, page_original_height_points)
        if result and isinstance(result, dict) and "products" in result and result["products"]:
            for product in result["products"]:
                if "bbox" in product:
                    product["bbox"] = normalize_bbox_to_pdf(product["bbox"], img.shape[0])
                product["page_num"] = page_num
            return result
    except Exception as e:
        logging.warning(f"Gemini OCR attempt failed for page {page_num}: {str(e)}. Falling back to Tesseract.")
    
    try:
        tesseract_result = tesseract_ocr_fallback(page_num, img, page_original_height_points)
        if tesseract_result and isinstance(tesseract_result, dict) and "products" in tesseract_result:
            return tesseract_result
        return None
    except Exception as e:
        logging.error(f"Both OCR methods failed for page {page_num}: {str(e)}")
        return None

# ... (keep all other existing functions the same)


# ENHANCED MATCHING LOGIC
# ======================
def match_products(ocr_item: Dict, master_df: pd.DataFrame) -> Dict:
    """Improved matching with multiple strategies"""
    ocr_id = str(ocr_item.get("id", "")).strip().lower()
    ocr_name = str(ocr_item.get("name", "")).strip().lower()
    ocr_category = str(ocr_item.get("category", "")).strip().lower()
    ocr_price = clean_price(ocr_item.get("price", ""))
    
    # Initialize result structure
    result = {
        "confidence": 0.0, # Use float for confidence
        "master_record": None,
        "ocr_data": ocr_item,
        "needs_review": True,
        "mismatched_fields": [],
        "match_strategy": "No Match" # For debugging/reporting
    }
    
    # 1. Exact ID match (flexible formatting)
    # Use `loc` for more efficient row selection and prevent chained assignment warnings
    exact_id_matches = master_df[master_df['商品CD'].astype(str).str.lower() == ocr_id]
    if not exact_id_matches.empty:
        result.update({
            "confidence": config.EXACT_MATCH_CONFIDENCE,
            "master_record": exact_id_matches.iloc[0].to_dict(),
            "needs_review": False,
            "match_strategy": "Exact ID Match"
        })
    
    if result["master_record"] is None:
        # 2. Name similarity with category filter
        if ocr_category:
            category_matches_df = master_df[
                master_df['大大分類 (家具／Hfa)'].astype(str).str.lower() == ocr_category
            ]
            if not category_matches_df.empty:
                # Use a list of choices for process.extractOne
                choices = category_matches_df['商品名'].astype(str).str.lower().tolist()
                if choices: # Ensure choices list is not empty
                    name_match_tuple = process.extractOne(
                        ocr_name,
                        choices,
                        scorer=fuzz.token_set_ratio,
                        score_cutoff=70
                    )
                    if name_match_tuple:
                        matched_name_str = name_match_tuple[0]
                        similarity_score = name_match_tuple[1]
                        
                        # Find the original row using the matched string
                        matched_row = category_matches_df[
                            category_matches_df['商品名'].astype(str).str.lower() == matched_name_str
                        ].iloc[0]
                        
                        result.update({
                            "confidence": similarity_score / 100.0 * 0.9, # Weight category match higher
                            "master_record": matched_row.to_dict(),
                            "match_strategy": "Category & Name Similarity"
                        })
    
    if result["master_record"] is None:
        # 3. Fallback to best name match across all products
        all_names_choices = master_df['商品名'].astype(str).str.lower().tolist()
        if all_names_choices: # Ensure choices list is not empty
            name_match_tuple = process.extractOne(
                ocr_name,
                all_names_choices,
                scorer=fuzz.token_set_ratio,
                score_cutoff=50
            )
            if name_match_tuple:
                matched_name_str = name_match_tuple[0]
                similarity_score = name_match_tuple[1]
                
                matched_row = master_df[
                    master_df['商品名'].astype(str).str.lower() == matched_name_str
                ].iloc[0]
                
                result.update({
                    "confidence": similarity_score / 100.0 * 0.7, # Lower weight for general name match
                    "master_record": matched_row.to_dict(),
                    "match_strategy": "General Name Similarity"
                })
    
    # Determine if review is needed based on the highest confidence achieved
    result["needs_review"] = result["confidence"] < config.GOOD_MATCH_CONFIDENCE
    
    # Check for field discrepancies if a master record was found
    if result["master_record"]:
        master_price = clean_price(result["master_record"].get("売価", ""))
        
        # Only compare prices if both are valid numbers
        if ocr_price is not None and master_price is not None:
            if abs(ocr_price - master_price) > config.PRICE_TOLERANCE:
                result["mismatched_fields"].append("price")
                result["needs_review"] = True # Flag for review if price mismatch
        
        # Check if OCR name is significantly different from master name for high-confidence matches
        if result["confidence"] >= config.GOOD_MATCH_CONFIDENCE and ocr_name and result["master_record"].get("商品名"):
            master_name_lower = str(result["master_record"]["商品名"]).strip().lower()
            if fuzz.ratio(ocr_name, master_name_lower) < 85: # Threshold for name similarity
                result["mismatched_fields"].append("name")
                result["needs_review"] = True # Flag for review if name mismatch

    # If still no match after all strategies, assign lowest confidence and mark for review
    if result["master_record"] is None:
        result.update({
            "confidence": 0.0,
            "needs_review": True,
            "match_strategy": "No Master Record Found"
        })
        logging.warning(f"No master record found for OCR item: ID='{ocr_id}', Name='{ocr_name}'.")

    return result

# ======================
# PDF ANNOTATION (GUARANTEED VISIBLE VERSION)
# ======================
def debug_page_coordinates(pdf_path: str):
    """Debug function to show page coordinates"""
    try:
        doc = fitz.open(pdf_path)
        for i in range(min(3, len(doc))):  # Check first 3 pages
            page = doc.load_page(i)
            logging.debug(f"\nPage {i+1} dimensions (width x height) in points: {page.rect.width} x {page.rect.height}")
            logging.debug(f"Page {i+1} media box: {page.mediabox}")
            logging.debug(f"Page {i+1} crop box: {page.cropbox}")
        doc.close()
    except Exception as e:
        logging.error(f"Failed to debug page coordinates for {pdf_path}: {e}")


def annotate_pdf_enhanced(pdf_path: str, matches: List[Dict], output_path: str) -> bool:
    """Enhanced PDF annotation with guaranteed visible markings"""
    try:
        # Open the source PDF
        doc = fitz.open(pdf_path)
        
        # Debug page coordinates first
        debug_page_coordinates(pdf_path)
        
        # Add a test annotation that should always be visible (on the first page)
        if doc.page_count > 0:
            try:
                test_page = doc.load_page(0)
                # Define a small rectangle in a corner, ensuring it's within page bounds
                test_rect_width = min(150.0, test_page.rect.width / 4)
                test_rect_height = min(50.0, test_page.rect.height / 10)
                test_rect = fitz.Rect(
                    test_page.rect.x1 - test_rect_width - 10,  # Right side, 10 points from edge
                    test_page.rect.y0 + 10,                     # Top, 10 points from edge
                    test_page.rect.x1 - 10,
                    test_page.rect.y0 + 10 + test_rect_height
                )
                test_shape = test_page.new_shape()
                test_shape.draw_rect(test_rect)
                test_shape.finish(color=(1, 0, 0), fill=None, width=3)  # Red border
                test_shape.commit()
                test_page.insert_text(
                    fitz.Point(test_rect.x0 + 5, test_rect.y0 + 15), # Inside the test rect
                    "TEST ANNOTATION - VISIBLE",
                    fontsize=8,
                    color=(1, 0, 0),
                    fontname="helv",
                    overlay=True
                )
                logging.info(f"Added test annotation on page 1 at {test_rect}")
            except Exception as e:
                logging.warning(f"Failed to add test annotation: {str(e)}")

        # Process each match
        for match in matches:
            ocr_data = match["ocr_data"]
            # Ensure bbox and page_num exist and are valid
            if not all(k in ocr_data for k in ["bbox", "page_num"]) or \
               not isinstance(ocr_data["bbox"], list) or len(ocr_data["bbox"]) != 4 or \
               not isinstance(ocr_data["page_num"], (int, float)):
                logging.warning(f"Skipping annotation for malformed OCR data: {ocr_data}")
                continue
                
            page_num_0_indexed = int(ocr_data["page_num"]) - 1 # Convert to 0-based index
            if page_num_0_indexed < 0 or page_num_0_indexed >= len(doc):
                logging.warning(f"Page number {ocr_data['page_num']} out of range for PDF {pdf_path}. Skipping annotation.")
                continue
                
            try:
                page = doc.load_page(page_num_0_indexed)
                bbox_coords = ocr_data["bbox"]
                
                # Convert to fitz.Rect and ensure validity
                # Clamp coordinates to page boundaries to avoid errors with out-of-bounds bboxes
                rect = fitz.Rect(
                    max(0.0, min(float(bbox_coords[0]), page.rect.width)),
                    max(0.0, min(float(bbox_coords[1]), page.rect.height)),
                    max(0.0, min(float(bbox_coords[2]), page.rect.width)),
                    max(0.0, min(float(bbox_coords[3]), page.rect.height))
                )
                
                # Ensure minimum size to make annotations visible even for tiny bboxes
                min_dim = 5.0 # Minimum dimension in points
                if rect.width < min_dim:
                    rect.x1 = rect.x0 + min_dim
                if rect.height < min_dim:
                    rect.y1 = rect.y0 + min_dim
                
                # Further ensure rect is valid and not empty
                if not rect.is_valid or rect.is_empty:
                    logging.warning(f"Invalid or empty rectangle for annotation on page {page_num_0_indexed+1}: {rect}. Original bbox: {bbox_coords}. Skipping.")
                    continue

                # Determine style based on confidence and review status
                if match["needs_review"]:
                    annotation_config = config.ANNOTATION_STYLES["needs_review"]
                elif match["confidence"] >= config.EXACT_MATCH_CONFIDENCE:
                     annotation_config = config.ANNOTATION_STYLES["high_confidence"]
                elif match["confidence"] >= config.GOOD_MATCH_CONFIDENCE:
                    annotation_config = config.ANNOTATION_STYLES["medium_confidence"]
                elif match["confidence"] >= config.MIN_MATCH_CONFIDENCE:
                    annotation_config = config.ANNOTATION_STYLES["low_confidence"]
                else: # Default if confidence falls below min_match_confidence
                    annotation_config = config.ANNOTATION_STYLES["low_confidence"] # Or define a 'no_match' style

                color = annotation_config.color
                thickness = annotation_config.thickness
                style = annotation_config.style
                
                # Create shape for guaranteed visibility
                shape = page.new_shape()
                
                if style == AnnotationStyle.BOX:
                    shape.draw_rect(rect)
                    shape.finish(color=color, fill=None, width=thickness)
                elif style == AnnotationStyle.HIGHLIGHT:
                    # PyMuPDF highlight is an annotation, not a shape.
                    # It's better to add a filled rectangle with opacity for "guaranteed visible"
                    # Using shape for consistency
                    shape.draw_rect(rect)
                    shape.finish(color=color, fill=color, width=0, fill_opacity=0.3)
                elif style == AnnotationStyle.UNDERLINE:
                    underline_rect = fitz.Rect(rect.x0, rect.y1 - thickness, rect.x1, rect.y1)
                    shape.draw_rect(underline_rect)
                    shape.finish(color=color, fill=color, width=0) # Use fill for solid underline
                elif style == AnnotationStyle.STRIKETHROUGH:
                    strikethrough_rect = fitz.Rect(rect.x0, rect.y0 + rect.height/2 - thickness/2, rect.x1, rect.y0 + rect.height/2 + thickness/2)
                    shape.draw_rect(strikethrough_rect)
                    shape.finish(color=color, fill=color, width=0)
                
                shape.commit()
                
                # Add text label (always visible above or below the box)
                label = f"ID: {ocr_data.get('id', 'N/A')} - Name: {ocr_data.get('name', 'N/A')[:30]}"
                if match["needs_review"]:
                    label += " (REVIEW)"
                
                # Position text label above the bbox to avoid overlapping content
                text_point = fitz.Point(rect.x0, rect.y0 - 5) # 5 points above the top of the rectangle
                
                # Ensure the text point is not going off the top of the page
                if text_point.y < 10: # If too close to top, place it below
                    text_point.y = rect.y1 + 12 

                page.insert_text(
                    text_point,
                    label,
                    fontsize=7, # Smaller font for more info
                    color=color,
                    fontname="helv",
                    overlay=True # Ensure text is visible
                )
                
                logging.info(f"Added {style.name} annotation for ID '{ocr_data.get('id', 'N/A')}' on page {page_num_0_indexed+1} at {rect}")
                
            except Exception as e:
                logging.error(f"Failed to annotate item on page {page_num_0_indexed+1} with OCR data {ocr_data.get('id', 'N/A')}: {str(e)}")
                continue
        
        # Save the document
        doc.save(output_path, garbage=4, deflate=True) # Optimize PDF size
        doc.close()
        logging.info(f"Successfully saved annotated PDF to {output_path}")
        return True
        
    except Exception as e:
        logging.error(f"PDF annotation failed completely for {pdf_path}: {str(e)}")
        return False

# ======================
# EXCEL LAYOUT REFERENCES
# ======================
def add_excel_layout_references(master_path: str, matches: List[Dict], source_pdf_filename: str) -> bool:
    """Add PDF layout references to Excel master"""
    try:
        # Read master file
        master_df = pd.read_excel(master_path)
        
        # Ensure all columns are present or add them
        tracking_cols_to_add = {
            "PDF_References": "",
            "PDF_Page": "",
            "PDF_BBox": "",
            "Confidence_Color": ""
        }
        for col, default_val in tracking_cols_to_add.items():
            if col not in master_df.columns:
                master_df[col] = default_val
        
        update_count = 0
        
        
        master_id_to_idx = {
            str(row['商品CD']).strip().lower(): idx
            for idx, row in master_df.iterrows()
        }

        for match in matches:
            ocr_data = match["ocr_data"]
            master_record = match["master_record"]
            
            if master_record is None: # Skip if no master record was matched
                continue

            master_id = str(master_record.get("商品CD", "")).strip().lower()
            
            if master_id not in master_id_to_idx:
                logging.warning(f"Master ID '{master_id}' from matched record not found in master_df index mapping. Skipping Excel update for this match.")
                continue

            idx = master_id_to_idx[master_id]
            
            # Update layout references
            master_df.at[idx, "PDF_References"] = source_pdf_filename
            master_df.at[idx, "PDF_Page"] = ocr_data.get("page_num", "")
            master_df.at[idx, "PDF_BBox"] = str(ocr_data.get("bbox", "")) # Store as string
            
            # Set confidence color
            if match["needs_review"]:
                color_hex = config.EXCEL_COLORS["needs_review"]
            elif match["confidence"] >= config.EXACT_MATCH_CONFIDENCE:
                color_hex = config.EXCEL_COLORS["high_confidence"]
            elif match["confidence"] >= config.GOOD_MATCH_CONFIDENCE:
                color_hex = config.EXCEL_COLORS["medium_confidence"]
            elif match["confidence"] >= config.MIN_MATCH_CONFIDENCE:
                color_hex = config.EXCEL_COLORS["low_confidence"]
            else:
                color_hex = config.EXCEL_COLORS["low_confidence"] # Default if confidence is very low
            
            master_df.at[idx, "Confidence_Color"] = color_hex
            update_count += 1
        
        # Save updated master
        # The file name should reflect that it's the original master with added data.
        # It's usually better to save a new file to avoid overwriting the original master.
        # Or, overwrite if that's the explicit requirement. Let's save a new one.
        base_name_master = os.path.splitext(os.path.basename(master_path))[0]
        updated_path = os.path.join(config.OUTPUT_DIR, f"{base_name_master}_with_layout_references.xlsx")
        
        # Create Excel writer with formatting engine
        writer = pd.ExcelWriter(updated_path, engine='xlsxwriter')
        master_df.to_excel(writer, sheet_name='Master Data', index=False)
        
        # Apply color formatting if requested
        if config.EXCEL_LAYOUT_REFERENCES:
            workbook = writer.book
            worksheet = writer.sheets['Master Data']
            
            # Create format objects based on hex colors
            format_dict = {}
            for name, hex_color in config.EXCEL_COLORS.items():
                format_dict[hex_color] = workbook.add_format({'bg_color': f'#{hex_color}'})
            
            # Get the column index for 'Confidence_Color'
            confidence_col_idx = master_df.columns.get_loc("Confidence_Color")
            
            # Apply formatting row by row based on the 'Confidence_Color' column
            # Start from row 2 because row 1 is headers (index 0) and data starts from index 1.
            for r_idx, row_data in master_df.iterrows():
                color_value = row_data.get('Confidence_Color')
                if color_value and color_value in format_dict:
                    # Apply format to the entire row + 1 (because excel rows are 1-indexed)
                    # and pandas dataframes are 0-indexed. Headers take first row.
                    worksheet.set_row(r_idx + 1, None, format_dict[color_value])
        
        writer.close()
        logging.info(f"Added layout references for {update_count} products in Excel file: {updated_path}")
        return True
        
    except Exception as e:
        logging.error(f"Excel layout reference update failed for {master_path}: {str(e)}")
        return False

# ======================
# MASTER FILE UPDATER
# ======================
def update_master_excel(master_path: str, matches: List[Dict], source_pdf_filename: str) -> bool:
    """Enhanced master Excel updater with smarter field updates"""
    try:
        master_df = pd.read_excel(master_path)
        
        # Add/ensure tracking columns exist
        tracking_cols = {
            'PDF_BBox': 'object', # Store string representation of bbox
            'Last_Updated': 'datetime64[ns]',
            'Source_PDF': 'object',
            'OCR_Confidence': 'float64',
            'OCR_Name': 'object',
            'OCR_Price': 'float64',
            'OCR_Category': 'object',
            'Needs_Review': 'bool',
            'Mismatched_Fields': 'object' # Store comma-separated string
        }
        
        for col, dtype in tracking_cols.items():
            if col not in master_df.columns:
                master_df[col] = pd.Series(dtype=dtype)
        
        update_count = 0
        
        # Create a dictionary for faster lookup of master records by 商品CD
        master_id_to_idx = {
            str(row['商品CD']).strip().lower(): idx
            for idx, row in master_df.iterrows()
        }

        for match in matches:
            ocr_data = match["ocr_data"]
            master_record = match["master_record"]
            
            if master_record is None:
                continue

            master_id = str(master_record.get("商品CD", "")).strip().lower()
            
            if master_id not in master_id_to_idx:
                continue

            idx = master_id_to_idx[master_id]
            
            # Update tracking fields
            master_df.at[idx, 'PDF_BBox'] = str(ocr_data.get('bbox', ''))
            master_df.at[idx, 'Last_Updated'] = datetime.now()
            master_df.at[idx, 'Source_PDF'] = source_pdf_filename
            master_df.at[idx, 'OCR_Confidence'] = match["confidence"]
            master_df.at[idx, 'Needs_Review'] = match.get("needs_review", False)
            master_df.at[idx, 'Mismatched_Fields'] = ", ".join(match.get("mismatched_fields", []))
            
            # Update OCR fields for reference
            master_df.at[idx, 'OCR_Name'] = ocr_data.get('name', '')
            master_df.at[idx, 'OCR_Price'] = clean_price(ocr_data.get('price', ''))
            master_df.at[idx, 'OCR_Category'] = ocr_data.get('category', '')
            
            # Conditionally update core fields based on config and confidence
            if config.UPDATE_CORE_FIELDS:
                # Logic for '商品名'
                if ('name' in ocr_data and ocr_data['name']):
                    current_master_name = master_df.at[idx, '商品名']
                    if pd.isna(current_master_name) or config.OVERWRITE_EXISTING or \
                       (match["confidence"] >= config.GOOD_MATCH_CONFIDENCE and \
                        fuzz.ratio(str(ocr_data['name']).lower(), str(current_master_name).lower()) > 85): # Only update if new OCR is better or overwrite
                        master_df.at[idx, '商品名'] = ocr_data['name']
                
                # Logic for '売価'
                ocr_price_cleaned = clean_price(ocr_data.get('price', ''))
                if ocr_price_cleaned is not None:
                    current_master_price = clean_price(master_df.at[idx, '売価'])
                    if pd.isna(current_master_price) or config.OVERWRITE_EXISTING or \
                       (match["confidence"] >= config.GOOD_MATCH_CONFIDENCE and \
                        abs(ocr_price_cleaned - current_master_price if current_master_price else 0) > 0): # Update if new price is different and confident
                        master_df.at[idx, '売価'] = ocr_price_cleaned
                
                # Logic for '大大分類 (家具／Hfa)'
                if 'category' in ocr_data and ocr_data['category']:
                    current_master_category = master_df.at[idx, '大大分類 (家具／Hfa)']
                    if pd.isna(current_master_category) or config.OVERWRITE_EXISTING or \
                       (match["confidence"] >= config.GOOD_MATCH_CONFIDENCE and \
                        str(ocr_data['category']).lower() != str(current_master_category).lower()):
                        master_df.at[idx, '大大分類 (家具／Hfa)'] = ocr_data['category']
            
            # Always update dimensions if provided by OCR
            if 'dimensions' in ocr_data and ocr_data['dimensions']:
                if 'Dimensions' not in master_df.columns:
                    master_df['Dimensions'] = None # Add new column if it doesn't exist
                master_df.at[idx, 'Dimensions'] = ocr_data['dimensions']
            
            update_count += 1
        
        # Save updated master
        base_name_master = os.path.splitext(os.path.basename(master_path))[0]
        updated_path = os.path.join(config.OUTPUT_DIR, f"{base_name_master}_updated_data.xlsx")
        master_df.to_excel(updated_path, index=False)
        logging.info(f"Updated {update_count} records in master Excel with OCR data: {updated_path}")
        return True
        
    except Exception as e:
        logging.error(f"Master update failed for {master_path}: {str(e)}")
        return False

# ======================
# MAIN WORKFLOW (UPDATED)
# ======================
def process_catalog(pdf_path: str, master_path: str) -> bool:
    """Complete processing with guaranteed visible annotations and structured outputs."""
    start_time = time.time()
    logging.info(f"Starting processing of {pdf_path}")
    print(f"\nStarting processing of {pdf_path}")
    
    try:
        # Enhanced input validation
        if not os.path.exists(pdf_path):
            logging.error(f"PDF file not found: {pdf_path}")
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")
        if not pdf_path.lower().endswith('.pdf'):
            logging.error(f"Input file must be a PDF: {pdf_path}")
            raise ValueError("Input file must be a PDF")
            
        if not os.path.exists(master_path):
            logging.error(f"Master file not found: {master_path}")
            raise FileNotFoundError(f"Master file not found: {master_path}")
        if not master_path.lower().endswith(('.xlsx', '.xls')):
            logging.error(f"Master file must be an Excel file: {master_path}")
            raise ValueError("Master file must be an Excel file")

        if not validate_master_file(master_path):
            logging.error("Master file validation failed. Aborting processing.")
            return False
            
        # Load master data
        master_df = pd.read_excel(master_path)
        logging.info(f"Loaded master data from {master_path} with {len(master_df)} rows.")
        
        # OCR Processing
        print("Converting PDF to images for OCR...")
        logging.info("Converting PDF to images for OCR...")
        pages_info = pdf_to_images(pdf_path) # Returns (page_num, img, pix_height, page_original_height_points)
        all_products_ocr_results = []
        
        # Process pages in parallel
        print(f"Performing OCR on {len(pages_info)} pages using {config.MAX_WORKERS} workers...")
        with ThreadPoolExecutor(max_workers=config.MAX_WORKERS) as executor:
            futures = []
            for page_num, img, pix_height, page_original_height_points in pages_info:
                # Submit OCR task for each page
                futures.append(executor.submit(
                    enhanced_ocr_extraction, 
                    page_num, 
                    img, 
                    page_original_height_points # Pass original page height for bbox scaling
                ))
            
            for future in futures:
                try:
                    ocr_result = future.result()
                    if ocr_result and isinstance(ocr_result, dict) and "products" in ocr_result:
                        all_products_ocr_results.extend(ocr_result["products"])
                except Exception as e:
                    logging.error(f"Error retrieving OCR result from future: {str(e)}")
                    continue
            
        logging.info(f"Found {len(all_products_ocr_results)} potential products from OCR.")
        print(f"Found {len(all_products_ocr_results)} potential products.")
        
        # Data Matching
        print("Matching OCR data with master records...")
        logging.info("Matching OCR data with master records...")
        matches = []
        for product_ocr_data in all_products_ocr_results:
            matches.append(match_products(product_ocr_data, master_df))
        
        logging.info(f"Completed matching. {len(matches)} matches processed.")
        
        # Generate Outputs
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        
        # Save JSON results
        json_path = os.path.join(config.OUTPUT_DIR, f"{base_name}_results.json")
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump({
                "source_pdf": os.path.basename(pdf_path),
                "processing_date": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                "products": matches,
                "stats": {
                    "total_items": len(matches),
                    "needs_review": sum(1 for m in matches if m.get("needs_review", False)),
                    "avg_confidence": sum(m.get("confidence", 0) for m in matches) / len(matches) if matches else 0
                }
            }, f, indent=2, default=str)
        logging.info(f"Saved JSON results to {json_path}")
        
        # Generate annotated PDF
        print("\nGenerating annotated PDF...")
        annotated_pdf_path = os.path.join(config.OUTPUT_DIR, f"{base_name}_annotated.pdf")
        if not annotate_pdf_enhanced(pdf_path, matches, annotated_pdf_path):
            logging.critical("Failed to create annotated PDF. Process will continue, but output will be incomplete.")
            # raise RuntimeError("Failed to create annotated PDF") # Decide if this is a critical failure
        
        # Add layout references to Excel
        print("Updating Excel layout references...")
        add_excel_layout_references(master_path, matches, os.path.basename(pdf_path))
        
        # Update master excel with OCR data
        print("Updating master Excel with OCR extracted data...")
        update_master_excel(master_path, matches, os.path.basename(pdf_path))
        
        # Generate review report if needed
        needs_review_items = [m for m in matches if m.get("needs_review", False)]
        if needs_review_items:
            review_df = pd.DataFrame([
                {
                    "Master_ID": m["master_record"]["商品CD"] if m["master_record"] else "N/A",
                    "Master_Name": m["master_record"]["商品名"] if m["master_record"] else "N/A",
                    "Master_Price": m["master_record"]["売価"] if m["master_record"] else "N/A",
                    "OCR_ID": m["ocr_data"].get("id", "N/A"),
                    "OCR_Name": m["ocr_data"].get("name", "N/A"),
                    "OCR_Price": m["ocr_data"].get("price", "N/A"),
                    "OCR_Category": m["ocr_data"].get("category", "N/A"),
                    "Page_Number": m["ocr_data"].get("page_num", "N/A"),
                    "Confidence": f"{m['confidence']:.2f}",
                    "Match_Strategy": m.get("match_strategy", "N/A"),
                    "Mismatched_Fields": ", ".join(m.get("mismatched_fields", [])),
                    "OCR_BBox": str(m["ocr_data"].get("bbox", "[]"))
                } for m in needs_review_items
            ])
            review_path = os.path.join(config.OUTPUT_DIR, f"{base_name}_needs_review_report.xlsx")
            review_df.to_excel(review_path, index=False)
            logging.info(f"Generated needs review report: {review_path}")
            print(f"Generated needs review report for {len(needs_review_items)} items: {review_path}")
        else:
            logging.info("No items flagged for review.")
            print("No items flagged for review.")
            
        end_time = time.time()
        duration = end_time - start_time
        logging.info(f"Processing completed for {pdf_path} in {duration:.2f} seconds.")
        print(f"\nProcessing completed successfully in {duration:.2f} seconds.")
        return True
        
    except Exception as e:
        logging.error(f"An unexpected error occurred during processing of {pdf_path}: {str(e)}", exc_info=True)
        print(f"\nERROR: Processing failed for {pdf_path}. Check logs for details. Error: {str(e)}")
        return False