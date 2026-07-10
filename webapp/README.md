# Bill → Excel Quote (Web App)

A one-page web app around your existing pipeline: upload a bill PDF, it runs
the Gemini dual-model extraction, fills the Excel template, and gives your
team a download button. No installs needed on their end — just a browser link.

## Why Streamlit Community Cloud (not Vercel)

You asked about Vercel — quick note on why this uses Streamlit instead:

- Your pipeline calls **two Gemini models with retry/backoff**, which can
  legitimately take 30–90+ seconds per bill (longer for multi-site
  consolidated quotes with several PDFs).
- Vercel's serverless functions default to a **10s timeout** and only go
  higher (60s on Hobby, up to a few minutes on Pro with Fluid Compute) with
  extra configuration — so slow bills would randomly fail with a 504 error.
- Streamlit Community Cloud runs your app as one long-lived Python process
  (like your laptop, but hosted) — no per-request timeout to fight, it's
  **free**, and it deploys straight from a GitHub repo with basically zero
  frontend code, since Streamlit *is* the frontend.

If you outgrow this later (e.g. want a custom-branded UI, or very high
traffic), the same pipeline code can be dropped behind a small FastAPI
backend on Render/Railway (which don't have Vercel's timeout ceiling either)
with a separate frontend — happy to build that version when you need it.

## What's in this folder

```
app.py                        <- the web app (Streamlit)
requirements.txt
single_gemini_client.py       <- your existing single-bill extraction
single_excel_filler.py        <- your existing single-bill Excel filler
consolidated_gemini_client.py <- your existing multi-site extraction
consolidated_excel_filler.py  <- your existing multi-site Excel filler
schema.py
template/                     <- your blank Excel templates (unchanged)
```

No pipeline logic was changed — `app.py` just calls the same functions your
CLI scripts already call.

## Deploy in ~5 minutes

1. **Create a GitHub repo** and push this whole folder to it.
   - Can be a **private** repo — Streamlit Community Cloud supports those.
   ```bash
   cd webapp
   git init
   git add .
   git commit -m "Bill to quote web app"
   git branch -M main
   git remote add origin https://github.com/<you>/bill-to-quote.git
   git push -u origin main
   ```

2. **Go to** [share.streamlit.io](https://share.streamlit.io) and sign in
   with GitHub.

3. Click **"New app"**, pick the repo/branch, set the main file to `app.py`,
   and deploy.

4. **Add your API key as a secret** (do this before or right after first
   deploy): in the app's **Settings → Secrets**, paste:
   ```toml
   GEMINI_API_KEY = "your-real-gemini-key"

   # Optional: gives the whole team a single shared password gate.
   # Leave this line out entirely if you don't want a password.
   APP_PASSWORD = "pick-something-simple"
   ```
   Save — the app restarts automatically with the key available.

5. You'll get a URL like `https://your-app-name.streamlit.app`. Share that
   with your team in the Philippines/India — works from any browser, no
   install.

## Using the app

- **Single site**: upload one bill PDF → download `quote_<filename>.xlsx`.
- **Consolidated**: upload one or more PDFs (a multi-site portfolio
  statement, or several separate site bills), optionally type the client
  name → download `consolidated_quote.xlsx`.
- If the two Gemini models disagree on any field, the app flags exactly
  which fields to double-check before sending the quote out — same
  cross-check logic as your CLI version.

## Local testing before you deploy

```bash
cd webapp
pip install -r requirements.txt
export GEMINI_API_KEY=your-real-key
streamlit run app.py
```
Opens at `http://localhost:8501`.

## Notes / things worth knowing

- **Secrets, not hardcoding**: the API key lives only in Streamlit's Secrets
  manager, never in the repo — safe even if the repo is later made public.
- **Cost**: Streamlit Community Cloud hosting is free for one app. Gemini
  API usage is billed the same way it already is in your CLI pipeline.
- **Concurrent use**: Community Cloud apps run on one shared instance — fine
  for a small team using it occasionally, but if several people upload big
  multi-site PDFs at the exact same time, requests will simply queue rather
  than run in parallel. Let me know if your team's volume grows and it's
  worth moving to a paid tier or a dedicated backend.
- **File size**: Streamlit's default upload limit is 200MB per file, far
  above any bill PDF, so no config needed there.
