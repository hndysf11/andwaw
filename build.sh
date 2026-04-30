#!/bin/bash
apt-get install -y poppler-utils tesseract-ocr tesseract-ocr-ara
pip install -r requirements.txt
python -m spacy download en_core_web_sm
python -m nltk.downloader punkt punkt_tab
