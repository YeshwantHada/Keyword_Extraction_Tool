# **Keyword Extraction Tool**

## **Overview**
This Python tool scans a directory (including subfolders) for multiple file types and extracts lines containing specified keywords. It supports **PDF, DOCX, RTF, TXT, and image files (TIFF/JPEG/JPG)** using OCR. The results are saved in an **Excel file** for easy review.

---

## **Features**
✔ Supports multiple file formats: PDF, DOCX, RTF, TXT, TIFF/JPEG/JPG  
✔ OCR for scanned images using Tesseract  
✔ Recursively scans directories  
✔ Saves results to Excel with file name, path, page number, and matching text  

---

## **Requirements**
Install dependencies:
```bash
pip install -r requirements.txt
```

`requirements.txt`:
```
PyMuPDF
python-docx
pytesseract
Pillow
pandas
striprtf
openpyxl
```

---

## **Usage**
1. Clone the repository:
```bash
git clone https://github.com/your-username/Keyword_Extraction_Tool.git
```
2. Navigate to the project folder:
```bash
cd multi-file-keyword-extractor
```
3. Update the `root_directory` variable in `src/Keyword_Extraction_Tool.py` to point to your folder containing files.
4. Run the script:
```bash
python src/Keyword_Extraction_Tool.py
```

---

## **Output**
The script creates an Excel file named `Output_results.xlsx` with columns:
- **File Name**
- **Full Path**
- **Page Number** (if applicable)
- **Matching Text**

**Example Output:**
| File Name      | Full Path                          | Page Number | Matching Text                     |
|---------------|-----------------------------------|-------------|----------------------------------|
| contract.pdf  | C:\Contracts\contract.pdf        | 3           | TERMINATION                     |
| doc1.docx     | C:\Contracts\doc1.docx          | None        | Service Level Exhibit           |

---

## **Customization**
- Modify the `keywords` list in the script to add or remove search terms:
```python
keywords = [
    "Service Level Exhibit",
    "Service Level Exhibit to Master Agreement",
    "TERMINATION"
]
```

---

## **License**
MIT License – Free to use and modify.

---

## **Author**
**Yeshwant Hada**  
https://github.com/YeshwantHada
