"""
CSV parser for bank-statement exports.

Auto-detects column layout across common formats (Revolut Business,
Bank of Cyprus, Hellenic Bank, generic debit/credit exports).
"""

from __future__ import annotations

import io
import logging
import re
from typing import Optional

import pandas as pd

from .config import (
    AMOUNT_ALIASES,
    CREDIT_ALIASES,
    CURRENCY_ALIASES,
    DATE_ALIASES,
    DEBIT_ALIASES,
    DESCRIPTION_ALIASES,
    FINAL_STATES,
    REFERENCE_ALIASES,
    STATE_ALIASES,
)
from .domain import BankTransaction
from .parsing import find_column_index, parse_date, parse_number


log = logging.getLogger(__name__)


class CsvParser:
    """Parses bank-statement CSV files into BankTransaction objects."""

    def parse(self, content: bytes) -> list[BankTransaction]:
        text = content.decode("utf-8-sig", errors="replace")
        separator = ";" if text.count(";") > text.count(",") else ","

        df = pd.read_csv(io.StringIO(text), sep=separator,
                         dtype=str, on_bad_lines="skip")
        df.columns = [c.strip() for c in df.columns]
        headers_lower = [c.lower() for c in df.columns]

        cols = self._resolve_columns(headers_lower)
        if cols["date"] == -1:
            log.warning("No date column found. Headers: %s", df.columns.tolist())
            return []

        transactions: list[BankTransaction] = []
        skipped_states: dict[str, int] = {}

        for idx, row in df.iterrows():
            cells = row.tolist()

            # Skip non-final transactions (e.g. PENDING, DECLINED, REVERTED)
            state = self._get_cell(cells, cols["state"]).upper()
            if cols["state"] != -1 and state and state not in FINAL_STATES:
                skipped_states[state] = skipped_states.get(state, 0) + 1
                continue

            txn = self._parse_row(cells, idx, cols)
            if txn is not None:
                transactions.append(txn)

        if skipped_states:
            log.info("Skipped non-final transactions: %s", skipped_states)

        return transactions

    # ── Column Resolution ─────────────────────────────────────────────────

    def _resolve_columns(self, headers_lower: list[str]) -> dict[str, int]:
        return {
            "date":        find_column_index(headers_lower, DATE_ALIASES),
            "description": find_column_index(headers_lower, DESCRIPTION_ALIASES),
            "reference":   find_column_index(headers_lower, REFERENCE_ALIASES),
            "amount":      find_column_index(headers_lower, AMOUNT_ALIASES),
            "debit":       find_column_index(headers_lower, DEBIT_ALIASES),
            "credit":      find_column_index(headers_lower, CREDIT_ALIASES),
            "currency":    find_column_index(headers_lower, CURRENCY_ALIASES),
            "state":       find_column_index(headers_lower, STATE_ALIASES),
        }

    # ── Row Parsing ───────────────────────────────────────────────────────

    def _parse_row(self, cells: list, idx: int, cols: dict[str, int]
                   ) -> Optional[BankTransaction]:

        date = parse_date(self._get_cell(cells, cols["date"]))
        if date is None:
            return None

        amount, txn_type = self._resolve_amount(
            cells, cols["amount"], cols["debit"], cols["credit"]
        )
        if amount == 0:
            return None

        currency = self._get_cell(cells, cols["currency"]) or "EUR"
        description = self._build_description(
            self._get_cell(cells, cols["description"]),
            self._get_cell(cells, cols["reference"]),
        )

        return BankTransaction(
            id=f"TXN-{idx + 1:06d}",
            date=date,
            description=description,
            amount=round(amount, 2),
            currency=currency,
            type=txn_type,
        )

    @staticmethod
    def _get_cell(cells: list, col_index: int) -> str:
        if col_index < 0 or col_index >= len(cells):
            return ""
        raw = cells[col_index]
        if raw is None or (isinstance(raw, float) and pd.isna(raw)):
            return ""
        return str(raw).strip()

    @staticmethod
    def _build_description(description: str, reference: str) -> str:
        """Combines description and reference fields intelligently."""
        description = description.strip()
        reference = reference.strip()

        # Skip references that are just technical IDs (all digits, UUIDs, etc.)
        if reference and re.fullmatch(r"[\d\-/]+", reference):
            reference = ""

        if description and reference and reference.lower() not in description.lower():
            return f"{description} · {reference}"
        if description:
            return description
        if reference:
            return reference
        return "—"

    def _resolve_amount(self, cells: list,
                        col_amount: int, col_debit: int, col_credit: int
                        ) -> tuple[float, str]:
        """Returns (absolute_amount, type) where type is 'Debit' or 'Credit'."""
        if col_amount != -1:
            value = parse_number(self._get_cell(cells, col_amount))
            if value is not None and value != 0:
                return abs(value), ("Debit" if value < 0 else "Credit")

        debit = parse_number(self._get_cell(cells, col_debit))
        credit = parse_number(self._get_cell(cells, col_credit))

        if debit and debit > 0:
            return debit, "Debit"
        if credit and credit > 0:
            return credit, "Credit"
        return 0, "Debit"
