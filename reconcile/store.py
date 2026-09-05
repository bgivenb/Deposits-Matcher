"""SQLite evidence store: atomic imports, immutable decisions, conflict-safe allocation."""

from collections import Counter, defaultdict
from contextlib import contextmanager
from dataclasses import asdict
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import sqlite3

from .domain import Transaction, canonical, digest, money, propose, read_statement

SCHEMA = """
CREATE TABLE IF NOT EXISTS imports(id TEXT PRIMARY KEY, source TEXT NOT NULL,
    filename TEXT NOT NULL, sha256 TEXT NOT NULL, row_count INTEGER NOT NULL);
CREATE TABLE IF NOT EXISTS transactions(uid TEXT PRIMARY KEY, payload TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS import_rows(import_id TEXT REFERENCES imports(id),
    row_number INTEGER NOT NULL, uid TEXT REFERENCES transactions(uid),
    PRIMARY KEY(import_id, row_number));
CREATE TABLE IF NOT EXISTS runs(id INTEGER PRIMARY KEY, created_at TEXT NOT NULL, report TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS candidates(id TEXT PRIMARY KEY, payload TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS run_candidates(run_id INTEGER REFERENCES runs(id),
    candidate_id TEXT REFERENCES candidates(id), PRIMARY KEY(run_id, candidate_id));
CREATE TABLE IF NOT EXISTS decisions(candidate_id TEXT PRIMARY KEY REFERENCES candidates(id),
    action TEXT NOT NULL CHECK(action IN ('accept','reject')), actor TEXT NOT NULL,
    reason TEXT NOT NULL, created_at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS allocations(uid TEXT PRIMARY KEY REFERENCES transactions(uid),
    candidate_id TEXT NOT NULL REFERENCES decisions(candidate_id));
CREATE TABLE IF NOT EXISTS audit(sequence INTEGER PRIMARY KEY, payload TEXT NOT NULL,
    previous_hash TEXT NOT NULL, hash TEXT NOT NULL);
CREATE TRIGGER IF NOT EXISTS audit_no_update BEFORE UPDATE ON audit BEGIN SELECT RAISE(ABORT,'audit is append-only'); END;
CREATE TRIGGER IF NOT EXISTS audit_no_delete BEFORE DELETE ON audit BEGIN SELECT RAISE(ABORT,'audit is append-only'); END;
CREATE TRIGGER IF NOT EXISTS decisions_no_update BEFORE UPDATE ON decisions BEGIN SELECT RAISE(ABORT,'decisions are immutable'); END;
CREATE TRIGGER IF NOT EXISTS decisions_no_delete BEFORE DELETE ON decisions BEGIN SELECT RAISE(ABORT,'decisions are immutable'); END;
CREATE TRIGGER IF NOT EXISTS transactions_no_update BEFORE UPDATE ON transactions BEGIN SELECT RAISE(ABORT,'imported records are immutable'); END;
CREATE TRIGGER IF NOT EXISTS transactions_no_delete BEFORE DELETE ON transactions BEGIN SELECT RAISE(ABORT,'imported records are immutable'); END;
CREATE TRIGGER IF NOT EXISTS candidates_no_update BEFORE UPDATE ON candidates BEGIN SELECT RAISE(ABORT,'candidates are immutable'); END;
CREATE TRIGGER IF NOT EXISTS candidates_no_delete BEFORE DELETE ON candidates BEGIN SELECT RAISE(ABORT,'candidates are immutable'); END;
"""


def now():
    return datetime.now(timezone.utc).isoformat(timespec="microseconds")


class Store:
    def __init__(self, filename):
        self.db = sqlite3.connect(filename, timeout=5, isolation_level=None)
        self.db.row_factory = sqlite3.Row
        self.db.execute("PRAGMA foreign_keys=ON")
        version = self.db.execute("PRAGMA user_version").fetchone()[0]
        if version not in (0, 1):
            self.db.close()
            raise ValueError("Workspace schema is newer than this application.")
        if (
            version == 0
            and self.db.execute(
                "SELECT 1 FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'"
            ).fetchone()
        ):
            self.db.close()
            raise ValueError("This is not an empty Deposits Matcher workspace.")
        self.db.execute("PRAGMA journal_mode=WAL")
        self.db.executescript(SCHEMA)
        self.db.execute("PRAGMA user_version=1")

    def close(self):
        self.db.close()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()

    @contextmanager
    def write(self):
        self.db.execute("BEGIN IMMEDIATE")
        try:
            self.verify_audit()
            yield
        except Exception:
            self.db.rollback()
            raise
        else:
            self.db.commit()

    def append_audit(self, event, actor, details):
        last = self.db.execute(
            "SELECT sequence, hash FROM audit ORDER BY sequence DESC LIMIT 1"
        ).fetchone()
        sequence, previous = (last[0] + 1, last[1]) if last else (1, "0" * 64)
        payload = canonical(
            {
                "sequence": sequence,
                "time": now(),
                "event": event,
                "actor": actor,
                "details": details,
            }
        )
        checksum = hashlib.sha256((previous + payload).encode()).hexdigest()
        self.db.execute(
            "INSERT INTO audit VALUES(?,?,?,?)", (sequence, payload, previous, checksum)
        )

    def verify_audit(self):
        previous, sequence = "0" * 64, 0
        for row in self.db.execute("SELECT * FROM audit ORDER BY sequence"):
            sequence += 1
            checksum = hashlib.sha256((previous + row["payload"]).encode()).hexdigest()
            if (
                row["sequence"] != sequence
                or row["previous_hash"] != previous
                or row["hash"] != checksum
            ):
                raise ValueError(f"Audit chain verification failed at sequence {sequence}.")
            previous = checksum
        return {"events": sequence, "head": previous}

    def import_statement(self, filename, source):
        with open(filename, "rb") as handle:
            content = handle.read(5_000_001)
        rows = read_statement(content, source)
        file_hash = hashlib.sha256(content).hexdigest()
        import_id = digest([source, file_hash])
        with self.write():
            existing = self.db.execute("SELECT id FROM imports WHERE id=?", (import_id,)).fetchone()
            if existing:
                return {"id": import_id, "inserted": 0, "existing": len(rows), "replayed": True}
            inserted = 0
            for row in rows:
                payload = canonical(asdict(row))
                prior = self.db.execute(
                    "SELECT payload FROM transactions WHERE uid=?", (row.uid,)
                ).fetchone()
                if prior and prior[0] != payload:
                    raise ValueError(
                        f"Conflicting existing {source} id {row.external_id!r} in account {row.account!r}; nothing was imported."
                    )
                if not prior:
                    self.db.execute("INSERT INTO transactions VALUES(?,?)", (row.uid, payload))
                    inserted += 1
            self.db.execute(
                "INSERT INTO imports VALUES(?,?,?,?,?)",
                (import_id, source, Path(filename).name, file_hash, len(rows)),
            )
            self.db.executemany(
                "INSERT INTO import_rows VALUES(?,?,?)",
                ((import_id, line, row.uid) for line, row in enumerate(rows, start=2)),
            )
            self.append_audit(
                "statement.imported",
                "local-import",
                {
                    "import_id": import_id,
                    "sha256": file_hash,
                    "source": source,
                    "inserted": inserted,
                    "rows": len(rows),
                },
            )
        return {
            "id": import_id,
            "inserted": inserted,
            "existing": len(rows) - inserted,
            "replayed": False,
        }

    def reconcile(self, window_days=3):
        with self.write():
            rows = [
                Transaction(**json.loads(row[0]))
                for row in self.db.execute(
                    "SELECT payload FROM transactions WHERE uid NOT IN (SELECT uid FROM allocations)"
                )
            ]
            if not rows:
                raise ValueError("No unresolved transactions. Import statements first.")
            report = propose(rows, window_days)
            # Rejected candidates stay rejected under the same versioned policy.
            decided = {row[0] for row in self.db.execute("SELECT candidate_id FROM decisions")}
            report["candidates"] = [
                item for item in report["candidates"] if item["id"] not in decided
            ]
            counts = Counter(
                uid for item in report["candidates"] for uid in item["bank"] + item["ledger"]
            )
            for item in report["candidates"]:
                item["ambiguous"] = any(counts[uid] > 1 for uid in item["bank"] + item["ledger"])
                # Ambiguity belongs to a run, not to a globally identified candidate.
                payload = {key: value for key, value in item.items() if key != "ambiguous"}
                self.db.execute(
                    "INSERT OR IGNORE INTO candidates VALUES(?,?)", (item["id"], canonical(payload))
                )
            cursor = self.db.execute(
                "INSERT INTO runs(created_at,report) VALUES(?,?)", (now(), canonical(report))
            )
            run_id = cursor.lastrowid
            self.db.executemany(
                "INSERT INTO run_candidates VALUES(?,?)",
                ((run_id, item["id"]) for item in report["candidates"]),
            )
            self.append_audit(
                "reconciliation.proposed",
                "local-engine",
                {
                    "run_id": run_id,
                    "input_digest": report["input_digest"],
                    "policy": report["policy"],
                    "window_days": window_days,
                    "candidates": len(report["candidates"]),
                    "limits": report["limits"],
                },
            )
        return {"run_id": run_id, **report}

    def decide(self, candidate_id, action, actor, reason):
        actor, reason = actor.strip(), reason.strip()
        if (
            action not in {"accept", "reject"}
            or not actor
            or len(actor) > 100
            or len(reason) < 8
            or len(reason) > 1000
        ):
            raise ValueError(
                "Decision requires accept/reject, an actor (1–100 characters), and a reason (8–1000 characters)."
            )
        if any(ord(char) < 32 for char in actor + reason):
            raise ValueError("Decision fields cannot contain control characters.")
        with self.write():
            row = self.db.execute(
                "SELECT payload FROM candidates WHERE id=?", (candidate_id,)
            ).fetchone()
            if not row:
                raise ValueError("Unknown candidate id; use the full id from the report.")
            prior = self.db.execute(
                "SELECT * FROM decisions WHERE candidate_id=?", (candidate_id,)
            ).fetchone()
            if prior:
                if (prior["action"], prior["actor"], prior["reason"]) == (action, actor, reason):
                    return {**dict(prior), "replayed": True}
                raise ValueError("Candidate already has an immutable decision.")
            candidate = json.loads(row[0])
            members = candidate["bank"] + candidate["ledger"]
            placeholders = ",".join("?" for _ in members)
            records = [
                Transaction(**json.loads(item[0]))
                for item in self.db.execute(
                    f"SELECT payload FROM transactions WHERE uid IN ({placeholders})", members
                )
            ]
            validated = propose(records, candidate["reason"]["date_window_days"])["candidates"]
            if not any(
                {key: value for key, value in item.items() if key != "ambiguous"} == candidate
                for item in validated
            ):
                raise ValueError("Candidate integrity check failed; no decision was recorded.")
            if self.db.execute(
                f"SELECT 1 FROM allocations WHERE uid IN ({placeholders})", members
            ).fetchone():
                raise ValueError(
                    "Candidate is stale: a transaction was already accepted in another match."
                )
            timestamp = now()
            self.db.execute(
                "INSERT INTO decisions VALUES(?,?,?,?,?)",
                (candidate_id, action, actor, reason, timestamp),
            )
            if action == "accept":
                self.db.executemany(
                    "INSERT INTO allocations VALUES(?,?)", ((uid, candidate_id) for uid in members)
                )
            self.append_audit(
                "candidate." + action,
                actor,
                {
                    "candidate_id": candidate_id,
                    "reason": reason,
                    "candidate_digest": digest(candidate),
                },
            )
        return {
            "candidate_id": candidate_id,
            "action": action,
            "actor": actor,
            "reason": reason,
            "created_at": timestamp,
            "replayed": False,
        }

    def report(self):
        self.db.execute("BEGIN")
        try:
            chain = self.verify_audit()
            latest = self.db.execute("SELECT * FROM runs ORDER BY id DESC LIMIT 1").fetchone()
            run = (
                {
                    "run_id": latest["id"],
                    "created_at": latest["created_at"],
                    **json.loads(latest["report"]),
                }
                if latest
                else None
            )
            decisions = {
                row["candidate_id"]: dict(row) for row in self.db.execute("SELECT * FROM decisions")
            }
            allocated = {
                row["uid"]: row["candidate_id"]
                for row in self.db.execute("SELECT * FROM allocations")
            }
            candidates, counts = [], Counter()
            for row in self.db.execute("SELECT * FROM candidates ORDER BY id"):
                item = json.loads(row["payload"])
                decision = decisions.get(item["id"])
                item["decision"] = decision
                item["status"] = (
                    {"accept": "accepted", "reject": "rejected"}[decision["action"]]
                    if decision
                    else (
                        "superseded"
                        if any(uid in allocated for uid in item["bank"] + item["ledger"])
                        else "pending"
                    )
                )
                if item["status"] == "pending":
                    counts.update(item["bank"] + item["ledger"])
                candidates.append(item)
            for item in candidates:
                item["ambiguous"] = item["status"] == "pending" and any(
                    counts[uid] > 1 for uid in item["bank"] + item["ledger"]
                )
            evaluated = set(run["input_ids"]) if run else set()
            rows, totals = (
                [],
                defaultdict(
                    lambda: {"bank": 0, "ledger": 0, "accepted_bank": 0, "accepted_ledger": 0}
                ),
            )
            for record in self.db.execute("SELECT payload FROM transactions ORDER BY uid"):
                row = json.loads(record[0])
                row["status"] = "accepted" if row["uid"] in allocated else "unresolved"
                row["accepted_candidate"] = allocated.get(row["uid"])
                row["pending_candidates"] = counts[row["uid"]]
                row["review_state"] = (
                    "accepted"
                    if row["uid"] in allocated
                    else "competing_candidates"
                    if counts[row["uid"]] > 1
                    else "awaiting_review"
                    if counts[row["uid"]]
                    else "not_evaluated"
                    if row["uid"] not in evaluated
                    else "search_limited"
                    if run and run["limits"]
                    else "no_pending_candidate"
                )
                rows.append(row)
                total = totals[(row["account"], row["currency"])]
                total[row["source"]] += row["cents"]
                if row["status"] == "accepted":
                    total["accepted_" + row["source"]] += row["cents"]
            summaries = [
                {
                    "account": account,
                    "currency": currency,
                    **{key: money(value) for key, value in total.items()},
                    "unresolved_bank": money(total["bank"] - total["accepted_bank"]),
                    "unresolved_ledger": money(total["ledger"] - total["accepted_ledger"]),
                }
                for (account, currency), total in sorted(totals.items())
            ]
            return {
                "schema_version": 1,
                "run": run,
                "imports": [
                    dict(row) for row in self.db.execute("SELECT * FROM imports ORDER BY id")
                ],
                "lineage": [
                    dict(row)
                    for row in self.db.execute(
                        "SELECT * FROM import_rows ORDER BY import_id,row_number"
                    )
                ],
                "transactions": rows,
                "candidates": candidates,
                "summary": summaries,
                "audit": chain,
                "audit_entries": [
                    dict(row) for row in self.db.execute("SELECT * FROM audit ORDER BY sequence")
                ],
            }
        finally:
            self.db.rollback()
