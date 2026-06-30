import sys
import fitz  # PyMuPDF
import pdfplumber

def extract_text(pdf_path):
    # Try PyMuPDF first
    try:
        doc = fitz.open(pdf_path)
        text = ""
        for page in doc:
            text += page.get_text()
        text = text.strip()
        if len(text) >= 50:
            return "PYMUPDF", text
    except Exception as e:
        sys.stderr.write(f"PyMuPDF failed: {str(e)}\n")

    # Fallback to pdfplumber
    try:
        text = ""
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        text = text.strip()
        if len(text) >= 50:
            return "PDFPLUMBER", text
    except Exception as e:
        sys.stderr.write(f"pdfplumber failed: {str(e)}\n")

    return "FAIL", ""

if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(1)
    pdf_path = sys.argv[1]
    method, text = extract_text(pdf_path)
    if method != "FAIL":
        print(f"---METHOD:{method}---")
        print(text)
    else:
        print("---METHOD:FAIL---")
