"""Decimal-safe matching logic for Deposits Matcher.

The UI is intentionally kept out of this module so the accounting behavior can be
tested without a display or Tkinter event loop.
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
import re
from typing import Iterable, Sequence


CENT = Decimal("0.01")
MAX_DEPOSITS_PER_LIST = 20


@dataclass(frozen=True)
class MatchResult:
    matched_indices_a: tuple[int, ...]
    matched_indices_b: tuple[int, ...]
    matched_total: Decimal
    unmatched_total_a: Decimal
    unmatched_total_b: Decimal


def parse_amount(value: str) -> Decimal:
    """Parse a non-negative currency value with at most two decimal places."""

    cleaned = value.strip()
    if cleaned.startswith("$"):
        cleaned = cleaned[1:].strip()
    if not cleaned:
        raise ValueError("Amount cannot be empty.")
    if not re.fullmatch(
        r"(?:\d{1,12}|\d{1,3}(?:,\d{3}){1,3})(?:\.\d{1,2})?", cleaned, flags=re.ASCII
    ):
        raise ValueError(
            "Use a non-negative amount, at most 12 integer digits and two decimal places."
        )
    cleaned = cleaned.replace(",", "")

    try:
        amount = Decimal(cleaned)
    except InvalidOperation as exc:
        raise ValueError(f"Invalid amount: {value!r}") from exc

    if not amount.is_finite():
        raise ValueError("Amount must be a finite number.")
    if amount < 0:
        raise ValueError("Amount cannot be negative.")
    if amount.quantize(CENT) != amount:
        raise ValueError("Amount cannot have more than two decimal places.")
    return amount.quantize(CENT)


def parse_amounts(values: Iterable[str]) -> list[Decimal]:
    return [parse_amount(value) for value in values]


def _to_cents(amount: Decimal) -> int:
    return int(amount / CENT)


def _subset_sums(amounts: Sequence[Decimal]) -> dict[int, int]:
    """Return one deterministic bit mask for each reachable total in cents."""

    sums: dict[int, int] = {0: 0}
    for index, amount in enumerate(amounts):
        cents = _to_cents(amount)
        additions = [(total + cents, mask | (1 << index)) for total, mask in sums.items()]
        for total, mask in additions:
            existing = sums.get(total)
            if existing is None or bin(mask).count("1") < bin(existing).count("1"):
                sums[total] = mask
    return sums


def _indices(mask: int, length: int) -> tuple[int, ...]:
    return tuple(index for index in range(length) if mask & (1 << index))


def find_maximum_match(amounts_a: Sequence[Decimal], amounts_b: Sequence[Decimal]) -> MatchResult:
    """Find the largest equal subset total without reusing an entry.

    Each list is limited to 20 entries because subset matching is exponential.
    The implementation stores values as integer cents, so comparisons are exact.
    """

    if len(amounts_a) > MAX_DEPOSITS_PER_LIST or len(amounts_b) > MAX_DEPOSITS_PER_LIST:
        raise ValueError(f"At most {MAX_DEPOSITS_PER_LIST} deposits are supported per list.")
    if any(
        not isinstance(amount, Decimal)
        or not amount.is_finite()
        or amount < 0
        or amount > Decimal("999999999999.99")
        or amount.quantize(CENT) != amount
        for amount in (*amounts_a, *amounts_b)
    ):
        raise ValueError("Amounts must be non-negative currency values in cents.")

    sums_a = _subset_sums(amounts_a)
    sums_b = _subset_sums(amounts_b)
    matched_cents = max(sums_a.keys() & sums_b.keys(), default=0)

    matched_total = Decimal(matched_cents) * CENT
    total_a = sum(amounts_a, start=Decimal("0.00"))
    total_b = sum(amounts_b, start=Decimal("0.00"))
    return MatchResult(
        matched_indices_a=_indices(sums_a[matched_cents], len(amounts_a)),
        matched_indices_b=_indices(sums_b[matched_cents], len(amounts_b)),
        matched_total=matched_total,
        unmatched_total_a=total_a - matched_total,
        unmatched_total_b=total_b - matched_total,
    )
