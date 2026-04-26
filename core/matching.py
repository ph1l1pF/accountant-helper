"""
Matching service: pairs each bank transaction with its best-matching receipt.

Uses a greedy strategy where each receipt can only be assigned to one
transaction. Transactions without a confident match get a list of best-guess
suggestions instead.
"""

from __future__ import annotations

from dataclasses import asdict
from datetime import datetime

from .config import (
    AMOUNT_MIN_TOLERANCE,
    AMOUNT_TOLERANCE_PERCENT,
    DAYS_TOLERANCE_HIGH,
    DAYS_TOLERANCE_MEDIUM,
    MAX_ALTERNATIVES,
)
from .domain import BankTransaction, Receipt


class MatchingService:
    """Pairs each transaction with its best-matching receipt."""

    CONFIDENCE_SCORE = {"High": 3, "Medium": 2, "Low": 1, "None": 0}

    # How close an amount must be (relative) to still be suggested as a best guess.
    SUGGESTION_MAX_RELATIVE_DIFF = 0.20     # 20 %
    SUGGESTION_MAX_ABSOLUTE_DIFF = 50.0     # or €50, whichever is larger
    SUGGESTION_LIMIT = 5

    def match(self, transactions: list[BankTransaction],
              receipts: list[Receipt]) -> list[dict]:
        assigned_receipt_names: set[str] = set()
        results: list[dict] = []

        for txn in transactions:
            ranked = self._rank_candidates(txn, receipts, assigned_receipt_names)

            best_receipt, best_confidence = (
                (ranked[0][0], ranked[0][1]) if ranked else (None, "None")
            )

            if best_receipt is not None:
                assigned_receipt_names.add(best_receipt.fileName)
                # For matched rows: alternatives = other receipts that also matched
                alternatives = [r for r, _ in ranked[1:1 + MAX_ALTERNATIVES]]
            else:
                # For unmatched rows: alternatives = best-guess suggestions
                alternatives = self._suggest_for_unmatched(
                    txn, receipts, assigned_receipt_names
                )

            results.append({
                "transaction": asdict(txn),
                "matchedReceipt": asdict(best_receipt) if best_receipt else None,
                "confidence": best_confidence,
                "alternativeCandidates": [asdict(r) for r in alternatives],
            })

        return results

    # ── Confident matches ─────────────────────────────────────────────────

    def _rank_candidates(self, txn: BankTransaction,
                         receipts: list[Receipt],
                         assigned: set[str]) -> list[tuple[Receipt, str]]:
        scored = [
            (r, self._calculate_confidence(txn, r))
            for r in receipts if r.fileName not in assigned
        ]
        scored = [(r, c) for r, c in scored if c != "None"]
        scored.sort(key=lambda pair: self.CONFIDENCE_SCORE[pair[1]], reverse=True)
        return scored

    def _calculate_confidence(self, txn: BankTransaction, receipt: Receipt) -> str:
        if receipt.extractedAmount is None:
            return "None"

        tolerance = max(txn.amount * AMOUNT_TOLERANCE_PERCENT, AMOUNT_MIN_TOLERANCE)
        if abs(txn.amount - receipt.extractedAmount) > tolerance:
            return "None"

        if not receipt.extractedDate:
            return "Low"

        days_delta = abs(
            (datetime.fromisoformat(txn.date) - datetime.fromisoformat(receipt.extractedDate)).days
        )
        if days_delta <= DAYS_TOLERANCE_HIGH:
            return "High"
        if days_delta <= DAYS_TOLERANCE_MEDIUM:
            return "Medium"
        return "Low"

    # ── Best-guess suggestions for unmatched transactions ─────────────────

    def _suggest_for_unmatched(self, txn: BankTransaction,
                               receipts: list[Receipt],
                               assigned: set[str]) -> list[Receipt]:
        """
        Returns up to SUGGESTION_LIMIT receipts ranked by combined similarity score.
        Considers receipts whose amount is within a relaxed tolerance (20 % or €50).
        """
        available = [
            r for r in receipts
            if r.fileName not in assigned and r.extractedAmount is not None
        ]

        scored: list[tuple[float, Receipt]] = []
        for receipt in available:
            amount_diff = abs(txn.amount - receipt.extractedAmount)
            relative_diff = amount_diff / max(txn.amount, 1.0)

            # Skip receipts that are clearly unrelated
            if (relative_diff > self.SUGGESTION_MAX_RELATIVE_DIFF
                    and amount_diff > self.SUGGESTION_MAX_ABSOLUTE_DIFF):
                continue

            score = self._suggestion_score(txn, receipt, amount_diff)
            scored.append((score, receipt))

        # Lower score = better match
        scored.sort(key=lambda pair: pair[0])
        return [r for _, r in scored[:self.SUGGESTION_LIMIT]]

    @staticmethod
    def _suggestion_score(txn: BankTransaction, receipt: Receipt,
                          amount_diff: float) -> float:
        """Combined score: amount difference weighted heavily, date proximity as tiebreaker."""
        amount_component = amount_diff  # euros off
        if receipt.extractedDate:
            days_delta = abs(
                (datetime.fromisoformat(txn.date)
                 - datetime.fromisoformat(receipt.extractedDate)).days
            )
            date_component = min(days_delta, 365) * 0.05
        else:
            date_component = 5  # small penalty for missing date
        return amount_component + date_component
