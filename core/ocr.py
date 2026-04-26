"""
OCR service: extracts amounts and dates from receipt files.

Uses pdfplumber for PDFs (fast, native text extraction) and Tesseract via
pytesseract for images. Images are downscaled and converted to grayscale
before OCR to bound memory usage on small servers.
"""

from __future__ import annotations

import io
import logging
import os
import re
from typing import Optional

import pdfplumber
import pytesseract
from PIL import Image

from .config import OCR_LANGUAGES
from .domain import Receipt
from .parsing import parse_date, parse_date_with_month_name, parse_number


log = logging.getLogger(__name__)


class OcrService:
    """Extracts amounts and dates from receipt files via Tesseract / pdfplumber."""

    # ── Amount patterns ────────────────────────────────────────────────
    # Matches a number with decimals, supporting both EU (1.234,56) and Anglo (1,234.56) formats.
    AMOUNT_NUMBER = r"[\d]{1,6}(?:[.,]\d{3})*[.,]\d{2}"

    # Priority 1 — "paid amount" keywords. Highest confidence — this is what hit the bank.
    # OCR-tolerant: Tesseract often reads "Bezahlter" as "Bezahiter" (l → i substitution).
    # The `\w{2,6}` catches any of: Bezahlter, Bezahiter, bezahlten, etc.
    PAID_KEYWORD_PATTERN = re.compile(
        r"(?:"
        r"bezah\w{2,6}\s+betrag"       # Bezahlter Betrag (and OCR variants)
        r"|zahlbetrag"
        r"|(?<![a-z])paid(?:\s+amount)?"
        r"|amount\s+(?:paid|charged)"
        r"|payment\s+(?:amount|received)"
        r"|total\s+(?:paid|charged)"
        r"|charged\s+(?:to|amount)"
        r")"
        r"\s*(?:\([^)]*\))?"           # optional parenthetical, e.g. "(EUR)"
        r"\s*[:\s]*"
        r"(?:EUR|€|USD|\$|GBP|£)?\s*"
        rf"({AMOUNT_NUMBER})",
        re.IGNORECASE,
    )

    # Priority 2 — "total" keywords. Covers English, German, Greek.
    TOTAL_KEYWORD_PATTERN = re.compile(
        r"(?:"
        r"(?<![a-z])total\s*(?:amount|due|payable)?"
        r"|grand\s*total"
        r"|amount\s*due"
        r"|balance\s*due"
        r"|gesamt(?:betrag)?"
        r"|endbetrag"
        r"|rechnungsbetrag"
        r"|(?<![a-zäöü])summe"
        r"|σύνολο"
        r")\b\s*[:\s]*"
        r"(?:EUR|€|USD|\$|GBP|£)?\s*"
        rf"({AMOUNT_NUMBER})",
        re.IGNORECASE,
    )

    # Priority 3 — any currency-marked amount (fallback, e.g. for 2-column layouts)
    CURRENCY_AMOUNT_PATTERN = re.compile(
        rf"(?:(?:EUR|€|USD|\$|GBP|£)\s*({AMOUNT_NUMBER})"
        rf"|({AMOUNT_NUMBER})\s*(?:EUR|€|USD|\$|GBP|£))",
        re.IGNORECASE,
    )

    # Keywords that mark an amount as HISTORICAL / superseded and should be ignored.
    # E.g. on Airbnb receipts: "Vorheriger Gesamtbetrag" (previous total before adjustment).
    STALE_AMOUNT_PREFIXES = ("vorherig", "previous", "alt", "old", "original")

    # ── Date patterns ──────────────────────────────────────────────────
    DATE_PATTERNS = [
        re.compile(r"\b(\d{4}-\d{2}-\d{2})\b"),
        re.compile(r"\b(\d{2}[./]\d{2}[./]\d{4})\b"),
    ]

    IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".tiff", ".tif", ".bmp"}

    # Cap the long side of input images before OCR to bound memory usage.
    # 2400px is still plenty for text recognition on receipts and keeps
    # each decoded image well under ~20 MB in RAM.
    MAX_IMAGE_DIMENSION = 2400

    AMOUNT_MIN = 0.5
    AMOUNT_MAX = 999_999

    def extract(self, content: bytes, filename: str) -> Receipt:
        ext = os.path.splitext(filename)[1].lower()
        text = self._extract_text(content, ext)

        return Receipt(
            fileName=filename,
            extractedAmount=self._extract_amount(text),
            extractedDate=self._extract_date(text),
        )

    # ── Text extraction ────────────────────────────────────────────────

    def _extract_text(self, content: bytes, extension: str) -> str:
        try:
            if extension == ".pdf":
                with pdfplumber.open(io.BytesIO(content)) as pdf:
                    return "\n".join(page.extract_text() or "" for page in pdf.pages)

            if extension in self.IMAGE_EXTENSIONS:
                with Image.open(io.BytesIO(content)) as image:
                    prepared = self._prepare_image_for_ocr(image)
                    try:
                        return pytesseract.image_to_string(prepared, lang=OCR_LANGUAGES)
                    finally:
                        prepared.close()

        except Exception as exc:
            log.warning("OCR failed for extension %s: %s", extension, exc)

        return ""

    def _prepare_image_for_ocr(self, image: "Image.Image") -> "Image.Image":
        """
        Down-scales very large images and converts to grayscale.
        Keeps memory usage bounded and — counter-intuitively — often
        improves OCR accuracy because Tesseract prefers modest resolutions
        with good contrast.
        """
        image.load()  # decode pixels now so we can safely close the source

        longest = max(image.size)
        if longest > self.MAX_IMAGE_DIMENSION:
            scale = self.MAX_IMAGE_DIMENSION / longest
            new_size = (int(image.size[0] * scale), int(image.size[1] * scale))
            resized = image.resize(new_size, Image.LANCZOS)
        else:
            resized = image.copy()

        # Grayscale cuts memory by ~3× vs RGB and doesn't hurt OCR on text.
        gray = resized.convert("L")
        if gray is not resized:
            resized.close()
        return gray

    # ── Amount extraction ──────────────────────────────────────────────

    def _extract_amount(self, text: str) -> Optional[float]:
        """
        Three-tier priority strategy:
        1. 'Bezahlter Betrag' / 'Paid' labels (what actually hit the bank — best match).
        2. 'Total' / 'Gesamtbetrag' labels, skipping 'Vorheriger' (historical) ones.
        3. Last currency-marked amount in the document (works for 2-column receipt layouts
           where the label text and amount appear in separate blocks).
        """
        if not text:
            return None

        paid = self._matches_with_filter(text, self.PAID_KEYWORD_PATTERN, group=1)
        if paid:
            return round(paid[-1], 2)

        totals = self._matches_with_filter(text, self.TOTAL_KEYWORD_PATTERN, group=1)
        if totals:
            return round(totals[-1], 2)

        all_currency = self._all_currency_amounts(text)
        if all_currency:
            return round(all_currency[-1], 2)

        return None

    def _matches_with_filter(self, text: str, pattern: re.Pattern, group: int) -> list[float]:
        """Returns parsed amounts for all matches, skipping those preceded by stale-amount words."""
        results: list[float] = []
        for match in pattern.finditer(text):
            # Check the 40 chars before the match for stale-amount words.
            context_before = text[max(0, match.start() - 40):match.start()].lower()
            if any(word in context_before for word in self.STALE_AMOUNT_PREFIXES):
                continue
            value = parse_number(match.group(group))
            if value is not None and self.AMOUNT_MIN < value < self.AMOUNT_MAX:
                results.append(value)
        return results

    def _all_currency_amounts(self, text: str) -> list[float]:
        """Extracts every currency-marked amount in document order."""
        values: list[float] = []
        for match in self.CURRENCY_AMOUNT_PATTERN.finditer(text):
            raw = match.group(1) or match.group(2)
            value = parse_number(raw)
            if value is not None and self.AMOUNT_MIN < value < self.AMOUNT_MAX:
                values.append(value)
        return values

    # ── Date extraction ────────────────────────────────────────────────

    def _extract_date(self, text: str) -> Optional[str]:
        if not text:
            return None

        # Numeric formats first (more reliable)
        for pattern in self.DATE_PATTERNS:
            match = pattern.search(text)
            if match:
                parsed = parse_date(match.group(1))
                if parsed:
                    return parsed

        # Fallback: German/English month names, e.g. "19. Februar 2026"
        return parse_date_with_month_name(text)
