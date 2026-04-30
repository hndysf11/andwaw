#!/bin/bash
set -e
pip install -r requirements.txt
python -m nltk.downloader punkt punkt_tab
