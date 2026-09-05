"""Strict statement ingestion and deterministic, bounded candidate generation.

Candidates are explanations for a reviewer, never accounting decisions.
"""

from collections import Counter, defaultdict
from dataclasses import asdict, dataclass
from datetime import date
import csv
import hashlib
from io import StringIO
from itertools import combinations
import json
import re

from matching import parse_amount

MAX_ROWS = 5_000
MAX_BYTES = 5_000_000
MAX_CANDIDATES = 10_000
MAX_COMPARISONS = 100_000
MAX_BATCH_POOL = 12
MAX_GROUP_SIZE = 4
POLICY_VERSION = "reference-date-v1"
REQUIRED = {"id", "date", "amount", "currency", "account", "reference"}


def canonical(value):
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False)


def digest(value):
    return hashlib.sha256(canonical(value).encode()).hexdigest()


def money(cents):
    return f"{cents // 100}.{cents % 100:02d}"


@dataclass(frozen=True)
class Transaction:
    uid: str
    source: str
    external_id: str
    date: str
    cents: int
    currency: str
    account: str
    reference: str
    description: str


def read_statement(content: bytes, source: str):
    """Validate the entire CSV before anything can be persisted."""
    if source not in {"bank", "ledger"}:
        raise ValueError("Source must be bank or ledger.")
    if len(content) > MAX_BYTES:
        raise ValueError("Statement exceeds the 5 MB limit.")
    try:
        text = content.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError("Statement must be UTF-8 CSV.") from exc
    reader = csv.DictReader(StringIO(text, newline=""), strict=True)
    try:
        headers = reader.fieldnames or []
    except csv.Error as exc:
        raise ValueError(f"Invalid CSV header: {exc}") from exc
    if len(set(headers)) != len(headers) or not REQUIRED.issubset(headers):
        raise ValueError(
            "Unique headers required: id,date,amount,currency,account,reference; description is optional."
        )
    if set(headers) - REQUIRED - {"description"}:
        raise ValueError("Unknown columns. Map your statement to the documented CSV schema first.")
    rows, seen = [], set()
    try:
        for line, raw in enumerate(reader, start=2):
            if len(rows) >= MAX_ROWS:
                raise ValueError(f"At most {MAX_ROWS} records per statement.")
            if None in raw or any(value is None for value in raw.values()):
                raise ValueError(f"Row {line}: incorrect column count.")
            row = {key: value.strip() for key, value in raw.items()}
            if any(
                len(value) > 512 or any(ord(char) < 32 for char in value) for value in row.values()
            ):
                raise ValueError(
                    f"Row {line}: fields must be at most 512 characters with no control characters."
                )
            if not row["id"] or not row["account"]:
                raise ValueError(f"Row {line}: id and account are required.")
            if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", row["date"], flags=re.ASCII):
                raise ValueError(f"Row {line}: date must be YYYY-MM-DD.")
            try:
                date.fromisoformat(row["date"])
                amount = parse_amount(row["amount"])
            except ValueError as exc:
                raise ValueError(f"Row {line}: {exc}") from exc
            if amount <= 0:
                raise ValueError(
                    f"Row {line}: deposit amounts must be positive; reversals are not supported."
                )
            currency = row["currency"]
            if not re.fullmatch(r"[A-Z]{3}", currency):
                raise ValueError(f"Row {line}: currency must be an uppercase three-letter code.")
            if row["amount"].startswith("$") and currency != "USD":
                raise ValueError(f"Row {line}: dollar-prefixed amounts require USD.")
            uid = digest([source, row["account"], row["id"]])
            if uid in seen:
                raise ValueError(f"Row {line}: duplicate id within this account and source.")
            seen.add(uid)
            rows.append(
                Transaction(
                    uid,
                    source,
                    row["id"],
                    row["date"],
                    int(amount * 100),
                    currency,
                    row["account"],
                    row["reference"],
                    row.get("description", ""),
                )
            )
    except csv.Error as exc:
        raise ValueError(f"Invalid CSV: {exc}") from exc
    if not rows:
        raise ValueError("Statement contains no transactions.")
    return rows


def propose(transactions, window_days=3):
    """Generate 1:1 and reference-backed 1:N/N:1 proposals, not a greedy allocation.

    Search is deterministic and resource-bounded. Explicit limits mean absence of
    a candidate never claims mathematical proof that no reconciliation exists.
    """
    if type(window_days) is not int or not 0 <= window_days <= 31:
        raise ValueError("Date window must be an integer from 0 to 31 days.")
    rows = sorted(transactions, key=lambda row: row.uid)
    if len(rows) > 2 * MAX_ROWS:
        raise ValueError(f"At most {2 * MAX_ROWS} unresolved transactions per run.")
    dates = {row.uid: date.fromisoformat(row.date).toordinal() for row in rows}
    amounts, references = defaultdict(list), defaultdict(list)
    for row in rows:
        amounts[(row.source, row.account, row.currency, row.cents)].append(row)
        if row.reference:
            references[(row.source, row.account, row.currency, row.reference)].append(row)
    found, limits = {}, set()
    comparisons = 0

    def spend():
        nonlocal comparisons
        comparisons += 1
        if comparisons > MAX_COMPARISONS or len(found) >= MAX_CANDIDATES:
            raise SearchLimit

    def add(bank, ledger, method):
        members = {
            "bank": sorted(row.uid for row in bank),
            "ledger": sorted(row.uid for row in ledger),
        }
        key = digest([POLICY_VERSION, window_days, members])
        found[key] = {
            "id": key,
            **members,
            "method": method,
            "cents": sum(row.cents for row in bank),
            "currency": bank[0].currency,
            "account": bank[0].account,
            "reason": {
                "date_window_days": window_days,
                "reference": bank[0].reference if method != "amount_date" else "",
                "exact_amount": True,
            },
        }

    try:
        for bank in (row for row in rows if row.source == "bank"):
            for ledger in amounts[("ledger", bank.account, bank.currency, bank.cents)]:
                spend()
                if abs(dates[bank.uid] - dates[ledger.uid]) > window_days:
                    continue
                if bank.reference and ledger.reference and bank.reference != ledger.reference:
                    continue
                add(
                    [bank],
                    [ledger],
                    "reference" if bank.reference and ledger.reference else "amount_date",
                )
        for anchor in rows:
            if not anchor.reference:
                continue
            other = "ledger" if anchor.source == "bank" else "bank"
            pool = []
            for row in references[(other, anchor.account, anchor.currency, anchor.reference)]:
                spend()
                if (
                    row.cents < anchor.cents
                    and abs(dates[row.uid] - dates[anchor.uid]) <= window_days
                ):
                    pool.append(row)
            if len(pool) > MAX_BATCH_POOL:
                limits.add(
                    f"Batch search skipped for {anchor.uid}: more than {MAX_BATCH_POOL} eligible records."
                )
                continue
            for size in range(2, min(MAX_GROUP_SIZE, len(pool)) + 1):
                for group in combinations(pool, size):
                    spend()
                    if sum(row.cents for row in group) == anchor.cents:
                        bank, ledger = (
                            ([anchor], group) if anchor.source == "bank" else (group, [anchor])
                        )
                        add(bank, ledger, "reference_batch")
    except SearchLimit:
        limits.add(
            "Search budget reached; proposals are partial. Narrow the input scope before drawing conclusions."
        )
    candidates = sorted(found.values(), key=lambda item: item["id"])
    participation = Counter(uid for item in candidates for uid in item["bank"] + item["ledger"])
    for item in candidates:
        item["ambiguous"] = any(participation[uid] > 1 for uid in item["bank"] + item["ledger"])
    return {
        "policy": POLICY_VERSION,
        "window_days": window_days,
        "candidates": candidates,
        "limits": sorted(limits),
        "comparisons": min(comparisons, MAX_COMPARISONS),
        "input_digest": digest([asdict(row) for row in rows]),
        "input_ids": [row.uid for row in rows],
    }


class SearchLimit(Exception):
    """Internal budget signal; never suppresses validation or storage errors."""
