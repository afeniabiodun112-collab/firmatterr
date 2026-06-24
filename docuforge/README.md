# DocuForge

Paste raw text → AI detects structure → Download a formatted `.docx`.

Powered by Google Gemini (gemini-3.1-flash-lite) + the `docx` npm package.

---

## Local Setup

```bash
cp .env.example .env
# Add your GEMINI_API_KEY to .env

npm install
npm start
# Open http://localhost:3000
```

Get a Gemini API key at https://aistudio.google.com/app/apikey

---

## Deploy to Render

1. Push this repo to GitHub
2. Go to https://render.com → New → Web Service
3. Connect your GitHub repo
4. Render auto-detects `render.yaml` — click **Apply**
5. In Environment Variables, set `GEMINI_API_KEY` to your Gemini key
6. Click **Deploy** — done. Your URL is live in ~60 seconds.

---

## What it does

| Feature | Detail |
|---|---|
| AI model | gemini-3.1-flash-lite via Gemini |
| Font | Times New Roman throughout |
| Title | 18pt bold centered |
| Headings | 14pt / 13pt bold |
| Body | 12pt, 1.15 line spacing |
| Margins | 1 inch all sides |
| Tables | Shaded header rows, auto-detected from prose |
| References | Auto-grouped at document end |

Files: `server.js` (all-in-one), `package.json`, `render.yaml`, `.env.example`
