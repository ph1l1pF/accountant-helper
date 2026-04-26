"""
Debug helper: run OCR on a single receipt file and print what was extracted.

Usage (from host):
    docker cp /path/to/receipt.jpg <container_id>:/tmp/receipt.jpg
    docker exec <container_id> python3 /app/debug_ocr.py /tmp/receipt.jpg
"""

import sys
from pathlib import Path

sys.path.insert(0, "/app")

from core.ocr import OcrService

if len(sys.argv) < 2:
    print("Usage: python3 debug_ocr.py <receipt_file>")
    sys.exit(1)

path = Path(sys.argv[1])
content = path.read_bytes()

svc = OcrService()
ext = path.suffix.lower()

print(f"\n{'='*60}")
print(f"File : {path.name}")
print(f"Size : {len(content)/1024:.1f} KB")
print(f"{'='*60}")

# Raw OCR text
text = svc._extract_text(content, ext)
print("\n── Raw OCR text ─────────────────────────────────────────")
print(text if text else "(empty)")

# Extracted fields
receipt = svc.extract(content, path.name)
print("\n── Extracted ────────────────────────────────────────────")
print(f"Amount : {receipt.extractedAmount}")
print(f"Date   : {receipt.extractedDate}")
print(f"{'='*60}\n")
