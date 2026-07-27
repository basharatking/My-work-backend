# RunDocs PDF Server v2.0

Handles all PDF processing tools for RunDocs.

## Run
```bash
pip install -r requirements.txt
uvicorn main:app --host 0.0.0.0 --port 8000
```

## Environment Variables
| Variable | Description |
|----------|-------------|
| `SUPABASE_URL` | Supabase project URL |
| `SUPABASE_SERVICE_KEY` | Supabase service role key |
| `PADDLE_WEBHOOK_SECRET` | Paddle webhook secret |

## Endpoints
- `GET /health` — Server status
- `POST /merge-pdf` — Merge PDFs
- `POST /split-pdf` — Split PDF
- `POST /compress-pdf` — Compress PDF
- `POST /rotate-pdf` — Rotate pages
- `POST /protect-pdf` — Password protect
- `POST /unlock-pdf` — Remove password
- `POST /add-watermark` — Add watermark
- `POST /add-page-numbers` — Add page numbers
- `POST /delete-pages` — Delete pages
- `POST /reorder-pages` — Reorder pages
- `POST /sign-pdf` — Sign PDF
- `POST /pdf-to-word` — PDF to Word
- `POST /pdf-to-excel` — PDF to Excel
- `POST /pdf-to-jpg` — PDF to JPG
- `POST /pdf-to-pptx` — PDF to PowerPoint
- `POST /pdf-to-text` — PDF to Text
- `POST /word-to-pdf` — Word to PDF
- `POST /excel-to-pdf` — Excel to PDF
- `POST /jpg-to-pdf` — Image to PDF
- `POST /html-to-pdf` — HTML to PDF
- `POST /pptx-to-pdf` — PowerPoint to PDF
- `POST /paddle/webhook` — Payment webhook
