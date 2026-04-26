"""
Configuration constants for the matching pipeline.

All values are env-driven where deployment may want to change them, and
hard-coded otherwise. This module has no dependencies on Flask or any other
web framework.
"""

from __future__ import annotations

import os
from pathlib import Path


# ── Server / runtime ─────────────────────────────────────────────────────

PORT = int(os.environ.get("PORT", 7860))
DEBUG_PORT: int | None = int(os.environ.get("DEBUG_PORT")) if os.environ.get("DEBUG_PORT") else None

# Where uploaded receipts are temporarily stored so the UI can preview them.
# Each match request creates a new session folder under UPLOAD_DIR.
UPLOAD_DIR = Path(os.environ.get("UPLOAD_DIR", "uploads")).resolve()
SESSION_TTL_HOURS = int(os.environ.get("SESSION_TTL_HOURS", 24))


# ── Matching tolerances ──────────────────────────────────────────────────

AMOUNT_TOLERANCE_PERCENT = 0.01   # 1 %
AMOUNT_MIN_TOLERANCE = 0.02       # €0.02 absolute minimum
DAYS_TOLERANCE_HIGH = 3
DAYS_TOLERANCE_MEDIUM = 14
MAX_ALTERNATIVES = 3


# ── OCR ──────────────────────────────────────────────────────────────────

# Tesseract language(s) – e.g. 'eng', 'eng+deu+ell'
OCR_LANGUAGES = os.environ.get("OCR_LANGS", "eng")


# ── CSV column aliases ───────────────────────────────────────────────────

# Order matters: first match wins, so put more-specific names first.
DATE_ALIASES = [
    # Revolut Business — prefer completion date over started date
    "date completed (utc)", "date completed",
    "date started (utc)",   "date started",
    # Other common bank formats
    "value date", "booking date", "transaction date", "posting date",
    "valuedate",  "bookingdate",  "transactiondate",
    "date", "datum",
]
DESCRIPTION_ALIASES = [
    "description", "details", "narrative", "memo",
    "particulars", "remarks", "merchant", "payee",
]
REFERENCE_ALIASES = ["reference", "ref", "note", "notes"]

AMOUNT_ALIASES   = ["amount", "value", "net amount"]
DEBIT_ALIASES    = ["debit", "withdrawal", "payments"]
CREDIT_ALIASES   = ["credit", "deposit", "receipts"]
CURRENCY_ALIASES = ["payment currency", "currency", "ccy"]
STATE_ALIASES    = ["state", "status"]

# Only these transaction states are considered final and matched.
# Used when a CSV has a state/status column (e.g. Revolut).
FINAL_STATES = {"COMPLETED", "DONE", "POSTED", "BOOKED", "SETTLED"}
