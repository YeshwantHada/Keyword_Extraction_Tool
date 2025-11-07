import os
import fitz  # PyMuPDF for PDF
import docx  # python-docx for DOCX
import pytesseract  # OCR for images
from PIL import Image
import pandas as pd
from striprtf.striprtf import rtf_to_text  # for RTF

# SLA-related phrases to search for
keywords = [
    "Service Level Exhibit",
    "Service Level Exhibit to Master Agreement",
	"TERMINATION"
]

# Supported file types
supported_extensions = ['.pdf', '.docx', '.tiff', '.tif', '.rtf', '.txt', '.jpeg', '.jpg']

# Extract text from PDF with page numbers
def extract_text_pdf(file_path):
    matches = []
    try:
        doc = fitz.open(file_path)
        for page_num, page in enumerate(doc, start=1):
            text = page.get_text()
            for line in text.splitlines():
                if any(keyword.lower() in line.lower() for keyword in keywords):
                    matches.append((page_num, line.strip()))
    except Exception as e:
        print(f"PDF error: {file_path} — {e}")
    return matches

# Extract text from DOCX
def extract_text_docx(file_path):
    matches = []
    try:
        doc = docx.Document(file_path)
        for para in doc.paragraphs:
            text = para.text.strip()
            if any(keyword.lower() in text.lower() for keyword in keywords):
                matches.append((None, text))
    except Exception as e:
        print(f"DOCX error: {file_path} — {e}")
    return matches

# Extract text from TIFF/JPEG/JPG using OCR
def extract_text_image(file_path):
    matches = []
    try:
        image = Image.open(file_path)
        text = pytesseract.image_to_string(image)
        for line in text.splitlines():
            if any(keyword.lower() in line.lower() for keyword in keywords):
                matches.append((None, line.strip()))
    except Exception as e:
        print(f"Image OCR error: {file_path} — {e}")
    return matches

# Extract text from RTF
def extract_text_rtf(file_path):
    matches = []
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            rtf_content = f.read()
        text = rtf_to_text(rtf_content)
        for line in text.splitlines():
            if any(keyword.lower() in line.lower() for keyword in keywords):
                matches.append((None, line.strip()))
    except Exception as e:
        print(f"RTF error: {file_path} — {e}")
    return matches

# Extract text from TXT
def extract_text_txt(file_path):
    matches = []
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                if any(keyword.lower() in line.lower() for keyword in keywords):
                    matches.append((None, line.strip()))
    except Exception as e:
        print(f"TXT error: {file_path} — {e}")
    return matches

# Unified text extractor
def extract_text(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.pdf':
        return extract_text_pdf(file_path)
    elif ext == '.docx':
        return extract_text_docx(file_path)
    elif ext in ['.tiff', '.tif', '.jpeg', '.jpg']:
        return extract_text_image(file_path)
    elif ext == '.rtf':
        return extract_text_rtf(file_path)
    elif ext == '.txt':
        return extract_text_txt(file_path)
    else:
        return []

# Scan folder recursively
def scan_directory(root_dir):
    results = []
    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            ext = os.path.splitext(filename)[1].lower()
            if ext in supported_extensions:
                full_path = os.path.join(dirpath, filename)
                print(f"Scanning: {full_path}")
                matches = extract_text(full_path)
                for page_num, match_text in matches:
                    results.append({
                        'File Name': filename,
                        'Full Path': full_path,
                        'Page Number': page_num,
                        'Matching Text': match_text
                    })
    return results

# Set your folder path here
root_directory = r"C:\Users\TEST\Testfolder" # change based on your repository

# Run scan
results = scan_directory(root_directory)

# Save to Excel
df = pd.DataFrame(results)
df.to_excel("Output_results.xlsx", index=False)

print(f"✅ Scan complete. {len(results)} matching entries saved to Output_results.xlsx.")
