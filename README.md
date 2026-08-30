# YT Title & Thumbnail Tool (Streamlit Cloud)

Extract YouTube titles and HD thumbnails for long-form videos.

## Deploy on Streamlit Community Cloud

1. Push these files to a **public GitHub** repository:
   - `app.py`
   - `requirements.txt`
   - `README.md` (optional)

2. Go to https://share.streamlit.io and sign in with GitHub.

3. **New app**
   - Repository: your repo
   - Branch: `main`
   - Main file path: `app.py`

4. Click **Deploy**. Your app URL will look like `https://xxxx.streamlit.app`.

## Local test (optional)

```bash
pip install -r requirements.txt
streamlit run app.py
```

This build always uses the cloud UI (browser downloads, no desktop folder picker).
