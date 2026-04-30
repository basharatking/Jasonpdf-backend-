# JasonPDF 🗂️
> **Free online PDF tools — named after Jason ❤️**  
> Merge · Split · Compress · Rotate · Watermark · Convert · Protect · Unlock

---

## 🗂️ Project Structure

```
jasonpdf/
├── frontend/           ← Static HTML/CSS/JS (deploy to Netlify / Vercel / GitHub Pages)
│   ├── index.html      ← Homepage with all tools grid
│   ├── shared.css      ← Global styles + dark mode
│   ├── shared.js       ← Shared logic (upload, API calls, nav, footer)
│   ├── config.js       ← ← EDIT THIS: set your backend URL
│   ├── about.html
│   ├── contact.html
│   ├── merge-pdf.html
│   ├── split-pdf.html
│   ├── compress-pdf.html
│   ├── rotate-pdf.html
│   ├── watermark-pdf.html
│   ├── protect-pdf.html
│   ├── unlock-pdf.html
│   ├── pdf-to-word.html
│   ├── pdf-to-excel.html
│   ├── pdf-to-jpg.html
│   ├── jpg-to-pdf.html
│   ├── pdf-to-pptx.html
│   ├── pdf-to-text.html
│   └── add-page-numbers.html
│
└── backend/            ← FastAPI Python backend
    ├── main.py         ← All 15 endpoints
    ├── requirements.txt
    ├── Procfile        ← For Railway / Heroku
    └── runtime.txt
```

---

## 🚀 Deployment Guide

### Step 1 — Deploy Backend (Railway — FREE)

1. Go to [railway.app](https://railway.app) → New Project → Deploy from GitHub
2. Push the `backend/` folder to a GitHub repo
3. Railway auto-detects the Procfile and deploys
4. Copy your Railway URL e.g. `https://jasonpdf-backend.up.railway.app`

**OR use Render.com:**
1. New Web Service → connect your GitHub repo
2. Build command: `pip install -r requirements.txt`
3. Start command: `uvicorn main:app --host 0.0.0.0 --port $PORT`

---

### Step 2 — Set Backend URL in Frontend

Open `frontend/config.js` and change:
```js
window.JASONPDF_CONFIG = {
  API_BASE: "https://YOUR-RAILWAY-URL.up.railway.app",  // ← paste here
  FREE_LIMIT_MB: 25,
};
```

---

### Step 3 — Deploy Frontend (Netlify — FREE)

1. Go to [netlify.com](https://netlify.com) → New Site
2. Drag and drop the entire `frontend/` folder
3. Done! Your site is live.

**OR GitHub Pages:**
1. Push `frontend/` to a GitHub repo
2. Settings → Pages → Source: main branch → `/` (root)

---

## 🔧 Local Development

### Backend
```bash
cd backend
pip install -r requirements.txt
uvicorn main:app --reload --port 8000
# API docs at: http://localhost:8000/docs
```

### Frontend
```bash
cd frontend
# Edit config.js → set API_BASE to http://localhost:8000
# Open index.html in browser or use live-server:
npx live-server
```

---

## 📋 All Tools (14 tools)

| Tool | Endpoint |
|------|----------|
| Merge PDF | `POST /merge-pdf` |
| Split PDF | `POST /split-pdf` |
| Compress PDF | `POST /compress-pdf` |
| Rotate PDF | `POST /rotate-pdf` |
| Watermark PDF | `POST /add-watermark` |
| PDF to Word | `POST /pdf-to-word` |
| PDF to Excel | `POST /pdf-to-excel` |
| PDF to JPG | `POST /pdf-to-jpg` |
| Image to PDF | `POST /jpg-to-pdf` |
| Unlock PDF | `POST /unlock-pdf` |
| Protect PDF | `POST /protect-pdf` |
| PDF to Text | `POST /pdf-to-text` |
| PDF to PowerPoint | `POST /pdf-to-pptx` |
| Add Page Numbers | `POST /add-page-numbers` |

---

## ⚙️ Environment Variables (Backend)

| Variable | Default | Description |
|----------|---------|-------------|
| `FREE_LIMIT_MB` | `25` | Max file size in MB |
| `PORT` | `8000` | Server port (auto-set by Railway/Render) |

---

## 🔒 Privacy

- Files are processed in memory and **deleted immediately** after response
- No user data is stored or logged
- No ads, no tracking

---

Built with ❤️ for Jason.
