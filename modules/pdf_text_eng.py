from pdf2image import convert_from_path
import pytesseract
import pandas as pd
import nltk
import sys
import os
from tqdm import tqdm

# Download English sentence tokenizer (first run only)
# Download English sentence tokenizer only if not already present
import nltk
try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    print("Downloading NLTK punkt tokenizer...")
    nltk.download('punkt')


# Replace with the actual path from `which tesseract` (macOS example)
import os
pytesseract.pytesseract.tesseract_cmd = os.environ.get("TESSERACT_PATH", "/usr/bin/tesseract")

def pdf_to_text_excel_english(pdf_path, output_txt, output_excel, lang="eng", progress_cb=None):
    pages = convert_from_path(pdf_path, poppler_path=os.environ.get("POPPLER_PATH", "/usr/bin"))
    all_text = ""

    # OCR each page
    for i, page in enumerate(tqdm(pages, desc="Processing pages {start}-{end}")):
        text = pytesseract.image_to_string(page, lang=lang)
        all_text += f"\n{text}\n"
        if progress_cb:
            progress_cb(i + 1, len(pages))

    # Save full text to .txt
    with open(output_txt, "w", encoding="utf-8") as f:
        f.write(all_text)
    print(f"✅ OCR text saved to {output_txt}")

    # Split text into sentences
    from nltk.tokenize import sent_tokenize
    sentences = sent_tokenize(all_text, language='english')

    # Save sentences to Excel
    df = pd.DataFrame(sentences, columns=['Sentence'])
    df.to_excel(output_excel, index=False)
    print(f"✅ Sentences saved to {output_excel}")

def main():
    if len(sys.argv) < 4:
        print("Usage: python pdf_to_txt.py input.pdf output.txt output.xlsx")
        sys.exit(1)

    pdf_file = sys.argv[1]
    output_txt = sys.argv[2]
    output_excel = sys.argv[3]

    if not os.path.exists(pdf_file):
        print(f"❌ Error: {pdf_file} not found.")
        sys.exit(1)

    pdf_to_text_excel_english(pdf_file, output_txt, output_excel)

if __name__ == "__main__":
    main()
