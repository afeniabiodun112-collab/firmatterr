# DocuForge

Paste raw text → AI detects structure → Download a formatted `.docx`.

Powered by Groq (llama-3.3-70b-versatile) + the `docx` npm package.

---

## Local Setup

```bash
cp .env.example .env
# Add your GROQ_API_KEY to .env

npm install
npm start
# Open http://localhost:3000
```

Get a free Groq API key at https://console.groq.com

---

## Deploy to Render

1. Push this repo to GitHub
2. Go to https://render.com → New → Web Service
3. Connect your GitHub repo
4. Render auto-detects `render.yaml` — click **Apply**
5. In Environment Variables, set `GROQ_API_KEY` to your Groq key
6. Click **Deploy** — done. Your URL is live in ~60 seconds.

---

## What it does

| Feature | Detail |
|---|---|
| AI model | llama-3.3-70b via Groq |
| Font | Times New Roman throughout |
| Title | 18pt bold centered |
| Headings | 14pt / 13pt bold |
| Body | 12pt, 1.15 line spacing |
| Margins | 1 inch all sides |
| Tables | Shaded header rows, auto-detected from prose |
| References | Auto-grouped at document end |

Files: `server.js` (all-in-one), `package.json`, `render.yaml`, `.env.example`
