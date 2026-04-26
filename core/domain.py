"""Domain dataclasses shared across the parsing, OCR, and matching layers."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass(frozen=True)
class BankTransaction:
    id: str
    date: str           # ISO yyyy-mm-dd
    description: str
    amount: float
    currency: str
    type: str           # 'Debit' | 'Credit'


@dataclass(frozen=True)
class Receipt:
    fileName: str
    extractedAmount: Optional[float]
    extractedDate: Optional[str]
