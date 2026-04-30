from pdf2image import convert_from_path
import pytesseract
import pandas as pd
from nltk.tokenize import sent_tokenize
import nltk
import sys
import os
from PyPDF2 import PdfReader
from tqdm import tqdm

# Path to tesseract executable
pytesseract.pytesseract.tesseract_cmd = "/usr/local/bin/tesseract"

# Use punkt from venv
nltk_data_path = os.path.join(os.path.dirname(sys.executable), "nltk_data")
if not os.path.exists(nltk_data_path):
    os.makedirs(nltk_data_path)

nltk.data.path.append(nltk_data_path)

try:
    nltk.data.find("tokenizers/punkt")
except LookupError:
    nltk.download("punkt", download_dir=nltk_data_path)

def pdf_to_text_excel(pdf_path, output_txt, output_excel, lang="ara", dpi=200, chunk_size=10, progress_cb=None):
    sentences = []

    total_pages = len(PdfReader(open(pdf_path, "rb")).pages)
    with open(output_txt, "w", encoding="utf-8") as f:
        pages_done = 0
        for start in range(1, total_pages + 1, chunk_size):
            end = min(start + chunk_size - 1, total_pages)
            pages = convert_from_path(
                pdf_path,
                dpi=dpi,
                first_page=start,
                last_page=end,
                poppler_path=os.environ.get("POPPLER_PATH", "/usr/bin")  # folder containing Poppler binaries
            )

            for page in tqdm(pages, desc=f"Processing pages {start}-{end}"):
                text = pytesseract.image_to_string(page, lang=lang)
                f.write(text + "\n")
                sentences.extend(sent_tokenize(text))
                pages_done += 1
                if progress_cb:
                    progress_cb(pages_done, total_pages)  # use sent_tokenize directly

    print(f"✅ OCR text saved to {output_txt}")

    df = pd.DataFrame(sentences, columns=["Sentence"])
    df.to_excel(output_excel, index=False)
    print(f"✅ Sentences saved to {output_excel}")

def main():
    if len(sys.argv) < 4:
        print("Usage: python pdf_to_txt_ar.py input.pdf output.txt output.xlsx")
        sys.exit(1)

    pdf_file = sys.argv[1]
    output_txt = sys.argv[2]
    output_excel = sys.argv[3]

    if not os.path.exists(pdf_file):
        print(f"❌ Error: {pdf_file} not found.")
        sys.exit(1)

    pdf_to_text_excel(pdf_file, output_txt, output_excel)

if __name__ == "__main__":
    main()
