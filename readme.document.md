# Imtihan Buddy (Groq edition) 📘⚡

Exam prep for Pakistani Matric and Intermediate students, built on the student's **own textbook chapter**.
Upload a chapter (PDF, DOCX, TXT or phone photo) and get board-style questions, answers checked against the chapter,
and explanations in English, اردو or Roman Urdu. Powered by the [Groq](https://console.groq.com) API.

## Features
- PDF / DOCX / TXT / phone-photo upload (photos read with a Groq vision model); edit the extracted text
- MCQs, short, long, fill-in-the-blank, true/false and numericals, each with a source reference
- Practice mode and timed Mock Test mode
- Rubric-based grading of typed or handwritten answers, grounded only in the chapter
- Trilingual explanations with RTL Urdu (Noto Nastaliq Urdu)
- Weak-topic breakdown, "Revise weak areas", bookmarks, optional saved history
- Chapter Q&A and summary, DOCX export of quiz + answer key

## Add your Groq API key
Get a free key at https://console.groq.com/keys, then use **any one** of these:
1. **Paste it in the app sidebar** ("Groq API key"). It is kept only for that browser session.
2. **Environment variable:** `export GROQ_API_KEY="gsk_..."` (PowerShell: `$env:GROQ_API_KEY="gsk_..."`)
3. **Secrets file:** create `.streamlit/secrets.toml` with `GROQ_API_KEY = "gsk_..."` (never commit this file)

## Run locally
```bash
git clone https://github.com/<your-username>/imtihan-buddy-groq.git
cd imtihan-buddy-groq
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud
Push these three files to GitHub, create the app with `app.py` as the main file, and add
`GROQ_API_KEY = "gsk_..."` under *Advanced settings > Secrets*.

## Models
Defaults are set at the top of the LLM section in `app.py`: text `llama-3.3-70b-versatile`, vision
`qwen/qwen3.6-27b` with `meta-llama/llama-4-scout-17b-16e-instruct` as fallback. Groq retires and adds models often, so
check https://console.groq.com/docs/models and override without editing code by setting `GROQ_MODEL` and/or
`GROQ_VISION_MODEL` (env var or secret).

## Free-tier notes
Groq's free tier has per-minute token limits. To stay inside them, the app sends at most ~6,000 characters of chapter
per request and generates questions in batches of 5. If you see the "server is busy" message, wait a few seconds and
tap again. `MAX_CALLS` caps requests per session. Urdu text uses more tokens than English.

## Privacy and copyright
Chapters stay in the session and are not stored. History is opt-in and keyed by a nickname (no login); on Community
Cloud the SQLite file is ephemeral. Use only material you have the right to study from.

## Recommended `.gitignore`
```
.venv/
__pycache__/
.streamlit/secrets.toml
imtihan.db
```
