# Master Students Builder — Lenient (Streamlit)

Upload Blackbaud, Rediker, and Student Records files → Download a styled Excel with:

- **Master** sheet: Source-colored text (BB black, RED dark red, SR dark green), yellow highlight for missing sources (by surname+grade).
- **Summary** sheet: ✅/❌ color-coded presence by source, with raw name strings.

## Setup

1. Create a **private GitHub repo** and upload these files.
2. Go to Streamlit Cloud → *New app* → choose this repo → select `app.py` → Deploy (Private).
3. In the app, upload your 3 source spreadsheets and click **Build Master Excel**.

## Local dev (optional)
```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```
