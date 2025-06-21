
from utilities import *

if __name__ == "__main__":
 
    """
    OPENAI_MODEL: "gpt-4-turbo"
    MAX_PAGES: 5
    PDF_DPI: 400
    IMAGE_PREPROCESS: true
    EXACT_MATCH_CONFIDENCE: 0.95
    GOOD_MATCH_CONFIDENCE: 0.85
    MIN_MATCH_CONFIDENCE: 0.4
    OVERWRITE_EXISTING: true
    ANNOTATION_STYLES:
      high_confidence:
        color: [0, 0.7, 0.1]
        thickness: 2.0
        style: BOX
      needs_review:
        color: [1, 0.2, 0.8]
        thickness: 2.5
        style: UNDERLINE
    TESSERACT_PATH: "/usr/local/bin/tesseract" # For macOS/Linux, adjust if needed
    """

    # --- Setup directories and dummy files for demonstration ---
    if not os.path.exists(config.INPUT_DIR):
        os.makedirs(config.INPUT_DIR)
    if not os.path.exists(config.OUTPUT_DIR):
        os.makedirs(config.OUTPUT_DIR)
    if not os.path.exists(config.DEBUG_DIR):
        os.makedirs(config.DEBUG_DIR)

    # Create a dummy master Excel file if it doesn't exist for testing
    master_excel_path = "product_maaster.xlsx"
    if not os.path.exists(master_excel_path):
        logging.info(f"Creating a dummy master Excel file at {master_excel_path}")
        dummy_data = {
            "商品CD": ["PROD001", "PROD002", "PROD003", "FURN001", "KIT005"],
            "商品名": ["Comfortable Sofa", "Modern Dining Table", "Elegant Chair", "Oak Wood Bookshelf", "Stainless Steel Pan"],
            "売価": [50000, 35000, 12000, 80000, 4500],
            "大大分類 (家具／Hfa)": ["furniture", "furniture", "furniture", "furniture", "kitchenware"],
            "ページ": [1, 2, 3, 1, 2],
            "ｸﾞﾘｯﾄﾞNo": ["A1", "B2", "C3", "D4", "E5"],
            "ｻｲｽﾞ": ["200x90x80", "150x80x75", "50x50x90", "180x30x200", "28cm"],
            "地域版": ["JP", "JP", "JP", "JP", "JP"],
            "対抗版": ["", "", "", "", ""],
            "掲載順": [1, 2, 3, 4, 5],
            "差替理由": ["", "", "", "", ""],
            "差替先": ["", "", "", "", ""],
            "小分類CD": ["SF", "DT", "CH", "BK", "KP"],
            "担当バイヤー": ["John Doe", "Jane Smith", "John Doe", "Jane Smith", "John Doe"],
            "商品名ｶﾅ": ["ソファ", "ダイニングテーブル", "イス", "本棚", "フライパン"],
            "掲載": ["Y", "Y", "Y", "Y", "Y"],
            "字組": ["", "", "", "", ""],
            "字組用商品名": ["", "", "", "", ""],
            "ｾｰﾙ売価": ["", "", "", "", ""],
            "売変開始日": ["", "", "", "", ""],
            "売変終了日": ["", "", "", "", ""],
            "材質名称": ["fabric", "wood", "plastic", "oak", "steel"],
            "色名称": ["grey", "brown", "white", "natural", "silver"],
            "幅": [200, 150, 50, 180, 28],
            "奥行": [90, 80, 50, 30, None],
            "高さ": [80, 75, 90, 200, None],
            "キャッチコピー": ["", "", "", "", ""],
            "枠付きコピー": ["", "", "", "", ""],
            "組立": ["Yes", "Yes", "No", "Yes", "No"],
            "注記": ["", "", "", "", ""],
            "セール情報備考": ["", "", "", "", ""],
            "連絡事項": ["", "", "", "", ""],
            "展開店舗割合": ["100%", "100%", "100%", "100%", "100%"],
            "ファイル名１": ["", "", "", "", ""], "ファイル名２": ["", "", "", "", ""],
            "ファイル名３": ["", "", "", "", ""], "ファイル名４": ["", "", "", "", ""], "ファイル名５": ["", "", "", "", ""],
            "ピクト１": ["", "", "", "", ""], "ピクト２": ["", "", "", "", ""], "ピクト３": ["", "", "", "", ""],
            "ピクト４": ["", "", "", "", ""], "ピクト５": ["", "", "", "", ""],
            "内寸": ["", "", "", "", ""], "矢印寸法": ["", "", "", "", ""]
        }
        dummy_df = pd.DataFrame(dummy_data)
        dummy_df.to_excel(master_excel_path, index=False)
        print(f"Dummy master.xlsx created at {master_excel_path}")

    # Process all PDFs in the input directory
    input_pdfs = [f for f in os.listdir(config.INPUT_DIR) if f.lower().endswith('.pdf')]

    if not input_pdfs:
        print(f"No PDF files found in '{config.INPUT_DIR}'. Please place your PDF catalogs there.")
        print("Example: Create a dummy PDF named 'sample_catalog.pdf' in 'input_pdfs'.")
    else:
        for pdf_file in input_pdfs:
            full_pdf_path = os.path.join(config.INPUT_DIR, pdf_file)
            print(f"\n--- Processing: {pdf_file} ---")
            success = process_catalog(full_pdf_path, master_excel_path)
            if success:
                print(f"Successfully processed {pdf_file}")
            else:
                print(f"Failed to process {pdf_file}")