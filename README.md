# Taxis & Logico-Semantic Shift Tool — Web Version

## Local setup

```bash
pip install -r requirements.txt
python -m spacy download en_core_web_sm
python -m nltk.downloader punkt punkt_tab
python app.py
```

Then open http://localhost:5000

## Deploy to Render (free hosting)

1. Push this folder to a GitHub repository
2. Go to https://render.com and sign up
3. Click "New" → "Web Service"
4. Connect your GitHub repo
5. Set these settings:
   - Build command: `pip install -r requirements.txt && python -m spacy download en_core_web_sm && python -m nltk.downloader punkt punkt_tab`
   - Start command: `gunicorn app:app --workers 2 --timeout 300 --bind 0.0.0.0:$PORT`
6. Click Deploy

Your tool will be live at a URL like: https://your-app-name.onrender.com

## Folder structure

```
taxis_web/
├── app.py              ← Flask backend
├── requirements.txt    ← Python dependencies
├── Procfile            ← Render start command
├── modules/            ← Copy your scripts here
│   ├── pdf_text_eng.py
│   ├── pdf_text_ar.py
│   ├── Eng_Analysis.py
│   ├── Ar_Analysis.py
│   └── Shift.py
└── templates/
    └── index.html      ← Frontend
```

## Important notes

- Copy your scripts from /Users/hind/Tool/ into the modules/ folder
- Tesseract and Poppler must be installed on the server for PDF OCR
- On Render free tier, the server sleeps after 15 min inactivity
  (first request after sleep takes ~30 seconds to wake up)
- For always-on hosting, upgrade to Render's paid plan ($7/month)
