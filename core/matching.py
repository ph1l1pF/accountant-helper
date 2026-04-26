"""
Matching service: pairs each bank transaction with its best-matching receipt.

Uses a globally-optimal greedy strategy:
  1. Score EVERY (transaction, receipt) pair up front.
  2. Sort all pairs globally — best score first.
  3. Assign greedily from that sorted list (both sides can only be matched once).

This avoids the classic failure of per-transaction-greedy where a transaction
processed early "steals" a receipt from a later transaction that would have
been a better fit (e.g. two transactions with the same amount where the one
with the slightly closer date should win).

Scoring uses three-level confidence (High/Medium/Low) as the primary sort key,
then days_diff and amount_diff as tiebreakers within each level.
"""

from __future__ import annotations

import math
from dataclasses import asdict, dataclass
from datetime import datetime
from typing import Optional

from .config import (
    AMOUNT_MIN_TOLERANCE,
    AMOUNT_TOLERANCE_PERCENT,
    DAYS_TOLERANCE_HIGH,
    DAYS_TOLERANCE_MEDIUM,
    MAX_ALTERNATIVES,
)
from .domain import BankTransaction, Receipt


# ── Internal score representation ────────────────────────────────────────

@dataclass(frozen=True)
class _Score:
    """Numeric representation of a (transaction, receipt) match quality."""
    tier: int           # 3 = High, 2 = Medium, 1 = Low
    amount_diff: float  # absolute € difference (lower is better)
    days_diff: float    # calendar days apart (inf when receipt has no date)

    @property
    def confidence(self) -> str:
        return {3: "High", 2: "Medium", 1: "Low"}[self.tier]

    def sort_key(self) -> tuple:
        """Higher tier first; within tier prefer smallest days then smallest amount."""
        return (-self.tier, self.days_diff, self.amount_diff)


# ── Matching service ──────────────────────────────────────────────────────

class MatchingService:
    """Pairs each transaction with its best-matching receipt."""

    # Relaxed tolerance for best-guess suggestions on unmatched transactions.
    SUGGESTION_MAX_RELATIVE_DIFF = 0.20     # 20 %
    SUGGESTION_MAX_ABSOLUTE_DIFF = 50.0     # or €50, whichever is larger
    SUGGESTION_LIMIT = 5

    def match(self, transactions: list[BankTransaction],
              receipts: list[Receipt]) -> list[dict]:
        # ── Phase 1: score every (txn, receipt) pair ──────────────────────
        scored_pairs: list[tuple[_Score, BankTransaction, Receipt]] = []
        for txn in transactions:
            for receipt in receipts:
                score = self._score(txn, receipt)
                if score is not None:
                    scored_pairs.append((score, txn, receipt))

        # ── Phase 2: globally optimal greedy assignment ────────────────────
        # Sort all pairs by quality so the best matches are assigned first,
        # regardless of the order transactions appear in the CSV.
        scored_pairs.sort(key=lambda p: p[0].sort_key())

        assigned: dict[str, tuple[Receipt, str]] = {}   # txn.id → (receipt, confidence)
        used_receipts: set[str] = set()

        for score, txn, receipt in scored_pairs:
            if txn.id not in assigned and receipt.fileName not in used_receipts:
                assigned[txn.id] = (receipt, score.confidence)
                used_receipts.add(receipt.fileName)

        # ── Phase 3: build result rows ────────────────────────────────────
        results: list[dict] = []
        for txn in transactions:
            if txn.id in assigned:
                matched_receipt, confidence = assigned[txn.id]

                # Alternatives = other receipts that also scored for this txn
                # but aren't assigned to anyone else.
                alternatives: list[Receipt] = []
                for _, t, r in scored_pairs:
                    if (t.id == txn.id
                            and r.fileName != matched_receipt.fileName
                            and r.fileName not in used_receipts):
                        alternatives.append(r)
                        if len(alternatives) >= MAX_ALTERNATIVES:
                            break
            else:
                matched_receipt = None
                confidence = "None"
                alternatives = self._suggest_for_unmatched(txn, receipts, used_receipts)

            results.append({
                "transaction": asdict(txn),
                "matchedReceipt": asdict(matched_receipt) if matched_receipt else None,
                "confidence": confidence,
                "alternativeCandidates": [asdict(r) for r in alternatives],
            })

        return results

    # ── Scoring ──────────────────────────────────────────────────────────

    def _score(self, txn: BankTransaction, receipt: Receipt) -> Optional[_Score]:
        """Returns a _Score if the receipt is a candidate for this transaction, else None."""
        if receipt.extractedAmount is None:
            return None

        tolerance = max(txn.amount * AMOUNT_TOLERANCE_PERCENT, AMOUNT_MIN_TOLERANCE)
        amount_diff = abs(txn.amount - receipt.extractedAmount)
        if amount_diff > tolerance:
            return None

        if not receipt.extractedDate:
            return _Score(tier=1, amount_diff=amount_diff, days_diff=math.inf)

        days_diff = abs(
            (datetime.fromisoformat(txn.date)
             - datetime.fromisoformat(receipt.extractedDate)).days
        )
        if days_diff <= DAYS_TOLERANCE_HIGH:
            tier = 3
        elif days_diff <= DAYS_TOLERANCE_MEDIUM:
            tier = 2
        else:
            tier = 1

        return _Score(tier=tier, amount_diff=amount_diff, days_diff=days_diff)

    # ── Best-guess suggestions for unmatched transactions ─────────────────

    def _suggest_for_unmatched(self, txn: BankTransaction,
                               receipts: list[Receipt],
                               used: set[str]) -> list[Receipt]:
        """
        Returns up to SUGGESTION_LIMIT receipts ranked by combined similarity score.
        Only considers receipts not already assigned to another transaction.
        """
        scored: list[tuple[float, Receipt]] = []
        for receipt in receipts:
            if receipt.fileName in used or receipt.extractedAmount is None:
                continue

            amount_diff = abs(txn.amount - receipt.extractedAmount)
            relative_diff = amount_diff / max(txn.amount, 1.0)

            if (relative_diff > self.SUGGESTION_MAX_RELATIVE_DIFF
                    and amount_diff > self.SUGGESTION_MAX_ABSOLUTE_DIFF):
                continue

            scored.append((self._suggestion_score(txn, receipt, amount_diff), receipt))

        scored.sort(key=lambda p: p[0])
        return [r for _, r in scored[:self.SUGGESTION_LIMIT]]

    @staticmethod
    def _suggestion_score(txn: BankTransaction, receipt: Receipt,
                          amount_diff: float) -> float:
        """Combined score: amount difference weighted heavily, date proximity as tiebreaker."""
        if receipt.extractedDate:
            days_delta = abs(
                (datetime.fromisoformat(txn.date)
                 - datetime.fromisoformat(receipt.extractedDate)).days
            )
            date_component = min(days_delta, 365) * 0.05
        else:
            date_component = 5
        return amount_diff + date_component
