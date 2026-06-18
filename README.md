# RunDocs — Improvements Installation Guide

⚠️ **No existing tools or pages were deleted.** Only fixes and additions.

---

## 📁 What's in `rundocs-frontend-fixes/`

Replace these files in your Netlify frontend repo (same filenames, just overwrite):

| File | What Changed |
|---|---|
| `word-to-pdf.html` | Now calls `/word-to-pdf` (was wrongly calling `/jpg-to-pdf`). Accepts `.doc/.docx` only. |
| `excel-to-pdf.html` | Now calls `/excel-to-pdf` (was wrongly calling `/jpg-to-pdf`). Accepts `.xls/.xlsx` only. |
| `delete-pages.html` | Now calls `/delete-pages` with a real page-number input (was wrongly calling `/split-pdf` with no input). |
| `reorder-pages.html` | Now calls `/reorder-pages` with a real page-order input (was wrongly calling `/split-pdf`). |
| `sign-pdf.html` | Now calls `/sign-pdf` with a real name/position form (was wrongly calling `/compress-pdf`). |
| `contact.html` | Added a WhatsApp "Get Pro" button + Pro plan feature card, so people can actually pay you. |
| `blog.html` | Replaced the 4 placeholder cards with 9 real SEO articles (bank statements, PDF to Excel, compress, AI tools, etc.) — each opens in a modal, links back to the matching tool. |

### Steps:
1. Open your GitHub repo for the frontend (or wherever you keep the Netlify files)
2. Replace the 7 files above with the new versions (same names — Netlify auto-redeploys)
3. **Before deploying:** open `contact.html` and replace the placeholder WhatsApp number `923001234567` with your real Pakistani WhatsApp Business number (format: `92` + number without the leading 0)

---

## 📁 What's in `rundocs-backend-additions/`

| File | Purpose |
|---|---|
| `additional_endpoints.py` | 5 new FastAPI endpoints: `/word-to-pdf`, `/excel-to-pdf`, `/delete-pages`, `/reorder-pages`, `/sign-pdf` |
| `requirements.txt` | Same as before + a note about LibreOffice |
| `replit.nix` | Tells Replit to install LibreOffice (needed for proper Word/Excel → PDF conversion) |

### Steps to install on your Replit backend:

1. **Add LibreOffice support:**
   - Open your Replit project
   - Create/edit the file `replit.nix` in the root and paste the contents from this package
   - This lets Replit's Nix package manager install LibreOffice automatically

2. **Add the new endpoints:**
   - Open `main.py` in your Replit backend
   - Open `additional_endpoints.py` from this package
   - Copy **everything below the imports comment block** and paste it into `main.py`, anywhere after the line `app = FastAPI(...)`
   - Add these two imports near the top of `main.py` (next to your other imports):
     ```python
     import subprocess
     import tempfile
     import shutil
     ```

3. **Restart your Repl** (Stop → Run). First boot after adding LibreOffice may take 1–2 minutes longer while Nix installs it.

4. **Test it:** visit `https://your-backend-url.replit.app/health` — should still return `{"status":"ok",...}`

### If LibreOffice fails to install on Replit:
The new endpoints have automatic fallbacks — `/word-to-pdf` and `/excel-to-pdf` will still work using a simplified text-only renderer (good enough for most cases, just won't preserve fancy formatting/images). `/delete-pages`, `/reorder-pages`, and `/sign-pdf` don't need LibreOffice at all — they work immediately.

---

## ✅ What This Fixes (Summary)

**Before:** 5 tool pages were silently calling the wrong backend endpoint — meaning Word→PDF would actually try to convert images, Delete Pages had no way to specify which pages, etc. Users got confusing or broken results.

**After:** Every tool calls its own correct, dedicated endpoint with proper input fields.

**Bonus:** Contact page can now actually convert visitors into paying Pro customers via WhatsApp, and the blog has real, Pakistan-relevant SEO content instead of 4 dummy cards with "Read More →" links that went nowhere.

---

## 🔜 Not Done Yet (for next round)
- Real payment gateway (currently WhatsApp manual orders — fine for low volume, but consider Stripe/Safepay/PayFast when volume grows)
- Migrating off Replit's free tier to Render/Railway for more reliable uptime
- Adding HBL/MCB/UBL-specific column parsing to the bank statement tool (currently uses the generic PDF-to-Excel engine)
