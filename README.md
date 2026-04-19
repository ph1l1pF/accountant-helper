# DocMatcher

Match bank transactions to receipt documents using OCR.
For tax accountants reconciling Cyprus Limited company bookkeeping.

## Quick start

### 1. Install Tesseract (OCR engine)

```bash
# macOS
brew install tesseract

# Ubuntu / Debian
sudo apt install tesseract-ocr

# Windows
# Download from: https://github.com/UB-Mannheim/tesseract/wiki
```

### 2. Install Python dependencies

```bash
pip install -r requirements.txt
```

### 3. Run

```bash
python server.py
```

Open **http://localhost:7860** in your browser.

---

## How it works

1. Upload your bank-statement CSV (column names auto-detected:
   `Date`, `Description`, `Amount` or `Debit`/`Credit`, `Currency`).
2. Select all receipt files at once (JPG, PNG, TIFF, PDF).
3. The server runs Tesseract on each receipt, extracts amount + date,
   then matches them to transactions.
4. Results show confidence: **High** (amount + date within 3 days),
   **Medium** (≤14 days), **Low** (amount only), **None** (no match).

## Configuration

Environment variables:

| Variable     | Default  | Purpose                                     |
|--------------|----------|---------------------------------------------|
| `PORT`       | `7860`   | HTTP port                                   |
| `OCR_LANGS`  | `eng`    | Tesseract languages, e.g. `eng+deu+ell`    |

To add German or Greek OCR:

```bash
# Install extra language packs
sudo apt install tesseract-ocr-deu tesseract-ocr-ell

# Run with multi-language OCR
OCR_LANGS=eng+deu+ell python server.py
```

## Architecture

```
docmatcher-py/
├── server.py            # Flask app with clean service classes:
│                        #   CsvParser, OcrService, MatchingService
├── requirements.txt
└── static/
    └── index.html       # Single-page frontend (vanilla JS)
```

## Next iteration ideas

- OneDrive integration (Microsoft Graph API)
- Manual override of assignments
- Export matched results to Excel
- Multi-language OCR for Greek receipts
