import os
import shutil
import re
import pytesseract
import pdf2image
import cv2
import numpy as np
from pdf2image import convert_from_path
from PyPDF2 import PdfReader
from PIL import Image

# Set Tesseract path (Windows users only, update accordingly)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Define the regex pattern
PATTERN = r"[A-Z]\d{7}"

# List to store names of PDFs with more than 20 pages
large_pdfs = []

def count_pdf_pages(pdf_path):
    """Counts the number of pages in a PDF file."""
    with open(pdf_path, "rb") as f:
        reader = PdfReader(f)
        return len(reader.pages)

def preprocess_image(image):
    """Enhances the image for better OCR recognition."""
    image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2GRAY)  # Convert to grayscale
    _, thresh = cv2.threshold(image, 150, 255, cv2.THRESH_BINARY_INV)  # Invert colors
    kernel = np.ones((2,2), np.uint8)
    processed_image = cv2.dilate(thresh, kernel, iterations=1)  # Enhance edges
    return processed_image

def extract_dates(text):
    """Extracts start and end dates from text, handling both YYYY/MM/DD and DD/MM/YYYY formats."""
    
    # Detect both YYYY/MM/DD and DD/MM/YYYY formats
    ymd_dates = re.findall(r"\b\d{4}/\d{2}/\d{2}\b", text)  # YYYY/MM/DD format
    dmy_dates = re.findall(r"\b\d{2}/\d{2}/\d{4}\b", text)  # DD/MM/YYYY format

    # Convert DD/MM/YYYY to YYYY/MM/DD
    converted_dates = [date for date in ymd_dates]  # Keep YYYY/MM/DD as is
    for date in dmy_dates:
        day, month, year = date.split('/')
        converted_dates.append(f"{year}/{month}/{day}")  # Convert to YYYY/MM/DD

    # Handle missing dates (handwritten detection)
    if not converted_dates:
        return (None, "HANDWRITTEN DATE - CHECK MANUALLY")

    return converted_dates[:2] if len(converted_dates) >= 2 else (converted_dates[0], None)

def find_string_in_pdf(pdf_path, search_pattern):
    """Searches for a string matching the given pattern in a PDF using OCR and extracts dates."""
    images = convert_from_path(pdf_path, poppler_path=r'C:\Users\Talenta Maluleke\Desktop\sortprog\poppler-24.08.0\Library\bin')
    
    for i, image in enumerate(images):
        processed_image = preprocess_image(image)  # Preprocess image for better OCR
        text = pytesseract.image_to_string(processed_image, config="--oem 1 --psm 6")
        print(f"\n--- Page {i+1} ---\n{text}")  # Print extracted text
        
        match = re.search(search_pattern, text)
        dates = extract_dates(text)
        
        if match:
            print(f"\n✅ String found: {match.group()}!")
            print(f"📅 Extracted Dates: Start: {dates[0]}, End: {dates[1]}")
            return match.group(), dates

    print("\n❌ String NOT found in extracted text.")
    return None, (None, "HANDWRITTEN DATE - CHECK MANUALLY")

def rename_and_move_pdf(original_pdf_path, matched_string):
    """Renames and moves the PDF to a new folder based on the matched string."""
    folder_name = matched_string.replace(' ', '_')
    os.makedirs(folder_name, exist_ok=True)
    
    new_pdf_name = f"{matched_string}.pdf"
    new_pdf_path = os.path.join(folder_name, new_pdf_name)
    
    if os.path.exists(new_pdf_path):
        print(f"File {new_pdf_path} already exists! Skipping move.")
        return
    
    shutil.move(original_pdf_path, new_pdf_path)
    print(f"Renamed and moved PDF to: {new_pdf_path}")
    
    return new_pdf_path  # Return new path for counting pages after moving

def process_pdfs_in_folder(folder_path, search_pattern):
    """Searches for a pattern in all PDFs in a folder and subfolders and moves matching files."""
    global large_pdfs
    
    for root, _, files in os.walk(folder_path):
        for file_name in files:
            if file_name.lower().endswith(".pdf"):
                pdf_path = os.path.join(root, file_name)
                
                # Count pages before processing
                page_count = count_pdf_pages(pdf_path)
                
                matched_string, dates = find_string_in_pdf(pdf_path, search_pattern)
                if matched_string:
                    new_pdf_path = rename_and_move_pdf(pdf_path, matched_string)
                    
                    # Count pages again after moving
                    if new_pdf_path:
                        moved_page_count = count_pdf_pages(new_pdf_path)
                        
                        if moved_page_count > 20:
                            large_pdfs.append((new_pdf_path, moved_page_count, dates[0], dates[1]))
                else:
                    print(f"Skipping {file_name}, no matching string found.")

    # Print list of PDFs with more than 20 pages
    print("\n📄 PDFs with more than 20 pages:")
    for pdf, pages, start_date, end_date in large_pdfs:
        print(f"{pdf} - Start Date: {start_date} - End Date: {end_date}")

# Example usage
if __name__ == "__main__":
    folder_to_scan = r"C:\Users\Talenta Maluleke\Desktop\sortprog"  # Change to actual folder path
    process_pdfs_in_folder(folder_to_scan, PATTERN)
