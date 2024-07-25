import os
import fitz  # PyMuPDF
from pdf2image import convert_from_path
from paddleocr import PaddleOCR
from docx import Document
import numpy as np
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
import argparse
import logging

# Set logging level
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Initialize PaddleOCR model
ocr = PaddleOCR(use_angle_cls=True, lang='ch')

def is_valid_content(text, min_valid_ratio=0.1):
    """Check if the text contains enough valid content"""
    total_chars = len(text)
    valid_chars = len(re.findall(r'\S', text))  # non-whitespace characters
    return valid_chars / total_chars > min_valid_ratio if total_chars > 0 else False

def convert_pdf_to_docx_with_ocr(pdf_path, docx_path):
    """Convert PDF to DOCX using OCR"""
    logging.info(f"Processing with OCR: {pdf_path}")
    doc = Document()
    try:
        for page in convert_from_path(pdf_path, dpi=300, first_page=1, last_page=None):
            image_np = np.array(page.convert('RGB'))
            result = ocr.ocr(image_np, cls=True)
            if result and result[0]:
                text_lines = [line[1][0] for line in result[0]]
                doc.add_paragraph("\n".join(text_lines))
            doc.add_page_break()
        doc.save(docx_path)
    except Exception as e:
        logging.error(f"Error processing PDF with OCR: {pdf_path}")
        logging.error(f"Error details: {str(e)}")

def convert_pdf_to_docx(pdf_path, docx_path):
    """Convert PDF to DOCX, use OCR if MuPDF error occurs"""
    try:
        doc = Document()
        with fitz.open(pdf_path) as pdf_document:
            for page in pdf_document:
                text = page.get_text("text")
                doc.add_paragraph(text)
                doc.add_page_break()
        doc.save(docx_path)
    except fitz.fitz.FileDataError as e:
        logging.warning(f"MuPDF error, switching to OCR: {pdf_path}")
        logging.warning(f"Error details: {str(e)}")
        convert_pdf_to_docx_with_ocr(pdf_path, docx_path)
    except Exception as e:
        logging.error(f"Error processing PDF: {pdf_path}")
        logging.error(f"Error details: {str(e)}")

def process_file(file_path, output_folder):
    """Process a single file"""
    filename = os.path.basename(file_path)
    docx_path = os.path.join(output_folder, os.path.splitext(filename)[0] + '.docx')
    
    if file_path.endswith('.pdf'):
        logging.info(f"Converting PDF: {filename}")
        convert_pdf_to_docx(file_path, docx_path)
        return file_path, docx_path
    else:
        return None

def process_files(input_folder, output_folder, max_workers=4):
    """Process all files in the folder"""
    processed_files = []
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_file = {}
        for filename in os.listdir(input_folder):
            file_path = os.path.join(input_folder, filename)
            if filename.endswith('.pdf'):
                future = executor.submit(process_file, file_path, output_folder)
                future_to_file[future] = file_path

        for future in as_completed(future_to_file):
            file_path = future_to_file[future]
            try:
                result = future.result()
                if result:
                    processed_files.append(result)
            except Exception as e:
                logging.error(f"Error processing file: {file_path}")
                logging.error(f"Error details: {str(e)}")
    
    return processed_files

def check_and_fix_files(processed_files, size_threshold):
    """Check and fix files based on size and content"""
    for pdf_path, docx_path in processed_files:
        need_ocr = False
        
        # Check file size
        if os.path.getsize(docx_path) < size_threshold:
            logging.warning(f"Found file smaller than {size_threshold/1024}KB: {docx_path}")
            need_ocr = True
        
        # Check file content
        if not need_ocr:
            doc = Document(docx_path)
            total_text = " ".join([paragraph.text for paragraph in doc.paragraphs])
            if not is_valid_content(total_text):
                logging.warning(f"File {docx_path} contains mainly blank content.")
                need_ocr = True
        
        # Reprocess with OCR if needed
        if need_ocr:
            logging.info(f"Reprocessing PDF with OCR: {pdf_path}")
            convert_pdf_to_docx_with_ocr(pdf_path, docx_path)
            logging.info(f"Reprocessed and saved: {docx_path}")

def main():
    parser = argparse.ArgumentParser(description="Convert PDF files to DOCX format")
    parser.add_argument("input_folder", help="Input folder path")
    parser.add_argument("output_folder", help="Output folder path")
    parser.add_argument("--max_workers", type=int, default=4, help="Maximum number of worker threads")
    parser.add_argument("--size_threshold", type=int, default=15, help="Small file threshold (KB)")
    args = parser.parse_args()

    if not os.path.exists(args.output_folder):
        os.makedirs(args.output_folder)

    processed_files = process_files(args.input_folder, args.output_folder, args.max_workers)
    logging.info("Checking and fixing files...")
    check_and_fix_files(processed_files, args.size_threshold * 1024)
    logging.info("All files processed.")

if __name__ == "__main__":
    main()