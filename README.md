# RunDocs Backend

FastAPI backend for RunDocs — PDF processing platform.

## Setup on Replit

1. Create a new Replit project (Python 3.11)
2. Upload all files from this folder
3. Add Secret: `HF_TOKEN` = your HuggingFace token (for AI tools)
4. Click Run — server starts on port 8000

## Environment Variables

| Variable   | Required | Description                          |
|------------|----------|--------------------------------------|
| HF_TOKEN   | Yes (AI) | HuggingFace API token for AI tools   |
| HF_MODEL   | No       | Override AI model (default: Mistral) |

## API Endpoints

### PDF Tools
- POST `/compress-pdf`      — Compress PDF
- POST `/merge-pdf`         — Merge multiple PDFs
- POST `/split-pdf`         — Split PDF pages
- POST `/rotate-pdf`        — Rotate PDF pages
- POST `/add-page-numbers`  — Add page numbers

### Convert
- POST `/pdf-to-word`       — PDF → DOCX
- POST `/pdf-to-excel`      — PDF → XLSX (smart grouping)
- POST `/pdf-to-jpg`        — PDF → JPG images (ZIP)
- POST `/pdf-to-pptx`       — PDF → PowerPoint
- POST `/pdf-to-text`       — PDF → plain text
- POST `/jpg-to-pdf`        — Images → PDF

### Security
- POST `/protect-pdf`       — Password protect PDF
- POST `/unlock-pdf`        — Remove PDF password
- POST `/add-watermark`     — Add text watermark

### AI Tools (requires HF_TOKEN)
- POST `/ai-summary`        — Summarize PDF
- POST `/ai-notes`          — Generate study notes
- POST `/ai-quiz`           — Generate quiz questions
- POST `/ai-keypoints`      — Extract key points
- POST `/ai-translate`      — Translate PDF content
- POST `/ask-pdf`           — Chat with PDF

### Utility
- POST `/ocr-check`         — Check if PDF is scanned
- GET  `/health`            — Health check

## PDF-to-Excel Smart Grouping

Tables with matching column headers across multiple pages are
automatically merged into a single Excel sheet. Different
table structures become separate sheets.

This fixes the old bug where each PDF page created a new sheet.
