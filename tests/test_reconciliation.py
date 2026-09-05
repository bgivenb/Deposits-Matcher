import csv
from dataclasses import replace
from io import StringIO
import json
from itertools import combinations
from pathlib import Path
import random
import sqlite3
import tempfile
import threading
import unittest
from unittest.mock import patch

from reconcile.__main__ import export_csv, main
from reconcile.domain import MAX_BATCH_POOL, Transaction, digest, propose, read_statement
from reconcile.store import Store

ROOT = Path(__file__).resolve().parents[1]


def statement(rows):
    handle = StringIO(newline="")
    writer = csv.writer(handle)
    writer.writerow(["id", "date", "amount", "currency", "account", "reference", "description"])
    for identifier, amount, reference in rows:
        writer.writerow(
            [identifier, "2026-08-03", amount, "USD", "operating", reference, "Synthetic test"]
        )
    return handle.getvalue().encode()


def txn(source, identifier, cents, reference="R", **changes):
    row = Transaction(
        digest([source, "operating", identifier]),
        source,
        identifier,
        "2026-08-03",
        cents,
        "USD",
        "operating",
        reference,
        "Synthetic",
    )
    return replace(row, **changes)


class IngestionTests(unittest.TestCase):
    def test_example_records_and_bom(self):
        content = (ROOT / "examples/bank.csv").read_bytes()
        self.assertEqual(len(read_statement(b"\xef\xbb\xbf" + content, "bank")), 10)

    def test_duplicate_ids_are_rejected_not_collapsed(self):
        with self.assertRaisesRegex(ValueError, "duplicate id"):
            read_statement(statement([("A", "10", ""), ("A", "10", "")]), "bank")

    def test_invalid_amounts_and_non_deposits(self):
        for amount in [
            "0",
            "-1",
            "NaN",
            "Infinity",
            "1e3",
            "1.001",
            "1,2",
            "1_000",
            "9999999999999",
            "1e999999",
        ]:
            with self.subTest(amount=amount), self.assertRaises(ValueError):
                read_statement(statement([("A", amount, "")]), "bank")

    def test_currency_date_and_field_validation(self):
        content = statement([("A", "10", "R")])
        for modified in [
            content.replace(b"USD", b"usd"),
            content.replace(b"2026-08-03", b"2026-02-30"),
            content.replace(b"2026-08-03", b"08/03/2026"),
            content.replace(b"operating", b""),
            content.replace(b"Synthetic test", b"=\tSUM(1)"),
        ]:
            with self.subTest(content=modified), self.assertRaises(ValueError):
                read_statement(modified, "bank")

    def test_invalid_headers_encoding_source_and_empty(self):
        for content, source in [
            (b"id,id\n1,2", "bank"),
            (b"\xff", "bank"),
            (statement([]), "bank"),
            (statement([("A", "10", "R")]), "other"),
        ]:
            with self.subTest(content=content), self.assertRaises(ValueError):
                read_statement(content, source)

    def test_caps_file_and_row_count(self):
        with patch("reconcile.domain.MAX_BYTES", 2), self.assertRaises(ValueError):
            read_statement(statement([("A", "10", "")]), "bank")
        with patch("reconcile.domain.MAX_ROWS", 1), self.assertRaises(ValueError):
            read_statement(statement([("A", "10", ""), ("B", "20", "")]), "bank")

    def test_explicit_dollar_prefix_is_usd_only(self):
        with self.assertRaises(ValueError):
            read_statement(statement([("A", "$10", "")]).replace(b"USD", b"EUR"), "bank")


class ProposalTests(unittest.TestCase):
    def test_exact_reference(self):
        result = propose([txn("bank", "A", 100), txn("ledger", "B", 100)])
        self.assertEqual(len(result["candidates"]), 1)
        self.assertEqual(result["candidates"][0]["method"], "reference")
        self.assertFalse(result["candidates"][0]["ambiguous"])

    def test_duplicate_amounts_are_explicitly_ambiguous(self):
        result = propose(
            [txn("bank", "A", 100, ""), txn("bank", "B", 100, ""), txn("ledger", "C", 100, "")]
        )
        self.assertEqual(len(result["candidates"]), 2)
        self.assertTrue(all(row["ambiguous"] for row in result["candidates"]))

    def test_reference_conflict_date_currency_account_boundaries(self):
        bank = txn("bank", "A", 100)
        for changes in [
            {"reference": "other"},
            {"date": "2026-08-07"},
            {"currency": "EUR"},
            {"account": "another"},
        ]:
            with self.subTest(changes=changes):
                self.assertEqual(
                    propose([bank, txn("ledger", "B", 100, **changes)])["candidates"], []
                )

    def test_date_window_is_inclusive(self):
        rows = [txn("bank", "A", 100), txn("ledger", "B", 100, date="2026-08-06")]
        self.assertEqual(len(propose(rows, 3)["candidates"]), 1)
        self.assertEqual(propose(rows, 2)["candidates"], [])

    def test_reference_backed_batch_in_both_directions(self):
        for left, right in [("bank", "ledger"), ("ledger", "bank")]:
            result = propose([txn(left, "A", 300), txn(right, "B", 100), txn(right, "C", 200)])
            self.assertEqual(len(result["candidates"]), 1)
            self.assertEqual(result["candidates"][0]["method"], "reference_batch")

    def test_batch_does_not_guess_without_reference(self):
        rows = [txn("bank", "A", 300, ""), txn("ledger", "B", 100, ""), txn("ledger", "C", 200, "")]
        self.assertEqual(propose(rows)["candidates"], [])

    def test_batch_pool_cap_is_visible_and_not_silently_truncated(self):
        rows = [txn("bank", "A", 1000)] + [
            txn("ledger", str(index), 100) for index in range(MAX_BATCH_POOL + 1)
        ]
        result = propose(rows)
        self.assertTrue(result["limits"])
        self.assertEqual(result["candidates"], [])

    def test_global_budget_is_reported(self):
        rows = [txn("bank", str(index), 100, "") for index in range(3)] + [
            txn("ledger", str(index), 100, "") for index in range(3)
        ]
        with patch("reconcile.domain.MAX_COMPARISONS", 2):
            result = propose(rows)
        self.assertEqual(result["comparisons"], 2)
        self.assertTrue(result["limits"])
        self.assertEqual(len(result["candidates"]), 2)

    def test_invalid_window(self):
        for value in [-1, 32, True, 1.5]:
            with self.subTest(value=value), self.assertRaises(ValueError):
                propose([], value)

    def test_order_independent_and_conservation_over_seeded_cases(self):
        randomizer = random.Random(17)
        for _ in range(50):
            rows = [
                txn(source, str(index), randomizer.randint(1, 15) * 100)
                for source in ["bank", "ledger"]
                for index in range(5)
            ]
            result = propose(rows)
            shuffled = list(rows)
            randomizer.shuffle(shuffled)
            self.assertEqual(result, propose(shuffled))
            by_id = {row.uid: row for row in rows}
            for item in result["candidates"]:
                self.assertEqual(
                    sum(by_id[uid].cents for uid in item["bank"]),
                    sum(by_id[uid].cents for uid in item["ledger"]),
                )
                self.assertEqual(
                    len(set(item["bank"] + item["ledger"])), len(item["bank"] + item["ledger"])
                )

    def test_indexed_unique_reference_workload(self):
        rows = [
            txn(source, str(index), index + 1, str(index))
            for source in ["bank", "ledger"]
            for index in range(1000)
        ]
        result = propose(rows)
        self.assertEqual(len(result["candidates"]), 1000)
        self.assertLess(result["comparisons"], 5000)
        self.assertFalse(result["limits"])

    def test_small_reference_batches_match_an_independent_brute_force_oracle(self):
        randomizer = random.Random(83)
        for _ in range(20):
            bank = [txn("bank", str(index), randomizer.randint(1, 12)) for index in range(4)]
            ledger = [txn("ledger", str(index), randomizer.randint(1, 12)) for index in range(4)]
            expected = set()
            for size_bank in range(1, 5):
                for size_ledger in range(1, 5):
                    if min(size_bank, size_ledger) != 1:
                        continue
                    for left in combinations(bank, size_bank):
                        for right in combinations(ledger, size_ledger):
                            if sum(row.cents for row in left) == sum(row.cents for row in right):
                                expected.add(
                                    (
                                        tuple(sorted(row.uid for row in left)),
                                        tuple(sorted(row.uid for row in right)),
                                    )
                                )
            actual = {
                (tuple(item["bank"]), tuple(item["ledger"]))
                for item in propose(bank + ledger)["candidates"]
            }
            self.assertEqual(actual, expected)


class StoreTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.filename = str(Path(self.temp.name) / "work.sqlite3")
        self.store = Store(self.filename)
        self.addCleanup(self.store.close)

    def ingest(self, source, rows, name=None):
        path = Path(self.temp.name) / (name or f"{source}.csv")
        path.write_bytes(statement(rows))
        return self.store.import_statement(path, source)

    def example(self):
        for source in ["bank", "ledger"]:
            self.store.import_statement(ROOT / f"examples/{source}.csv", source)
        return self.store.reconcile()

    def test_import_replay_and_overlapping_files_preserve_identity(self):
        self.ingest("bank", [("A", "10", "R")])
        replay = self.ingest("bank", [("A", "10", "R")])
        self.assertTrue(replay["replayed"])
        self.assertEqual(self.store.verify_audit()["events"], 1)
        overlap = self.ingest("bank", [("A", "10", "R"), ("B", "20", "S")], "later.csv")
        self.assertEqual((overlap["inserted"], overlap["existing"]), (1, 1))
        self.assertEqual(len(self.store.report()["transactions"]), 2)

    def test_conflicting_import_rolls_back_earlier_new_rows(self):
        self.ingest("bank", [("A", "10", "R")])
        with self.assertRaisesRegex(ValueError, "Conflicting"):
            self.ingest("bank", [("B", "20", "S"), ("A", "99", "R")])
        self.assertEqual(len(self.store.report()["transactions"]), 1)
        self.assertEqual(self.store.verify_audit()["events"], 1)

    def test_same_external_id_is_scoped_to_source(self):
        for source in ["bank", "ledger"]:
            self.ingest(source, [("A", "10", "R")])
        self.assertEqual(len(self.store.report()["transactions"]), 2)

    def test_new_import_after_a_run_is_not_falsely_reported_as_evaluated(self):
        self.example()
        self.ingest("bank", [("NEW", "999", "NEW-REFERENCE")], "new.csv")
        row = next(
            item for item in self.store.report()["transactions"] if item["external_id"] == "NEW"
        )
        self.assertEqual(row["review_state"], "not_evaluated")
        self.assertEqual(row["status"], "unresolved")

    def test_example_proposals_are_not_auto_accepted(self):
        result = self.example()
        self.assertEqual(len(result["candidates"]), 6)
        self.assertEqual(sum(item["ambiguous"] for item in result["candidates"]), 2)
        self.assertTrue(
            all(row["status"] == "unresolved" for row in self.store.report()["transactions"])
        )

    def test_acceptance_replay_and_conflict(self):
        result = self.example()
        choices = [item for item in result["candidates"] if item["ambiguous"]]
        accepted = self.store.decide(
            choices[0]["id"], "accept", "Reviewer", "Confirmed against source evidence"
        )
        replay = self.store.decide(
            choices[0]["id"], "accept", "Reviewer", "Confirmed against source evidence"
        )
        self.assertFalse(accepted["replayed"])
        self.assertTrue(replay["replayed"])
        with self.assertRaisesRegex(ValueError, "stale"):
            self.store.decide(
                choices[1]["id"], "accept", "Other", "Confirmed against source evidence"
            )
        report = self.store.report()
        self.assertEqual(sum(row["status"] == "accepted" for row in report["transactions"]), 2)
        self.assertEqual(sum(row["status"] == "superseded" for row in report["candidates"]), 1)

    def test_rejection_persists_and_does_not_allocate(self):
        candidate = self.example()["candidates"][0]
        self.store.decide(candidate["id"], "reject", "Reviewer", "Insufficient source evidence")
        self.assertTrue(
            all(row["status"] == "unresolved" for row in self.store.report()["transactions"])
        )
        self.assertNotIn(
            candidate["id"], {item["id"] for item in self.store.reconcile()["candidates"]}
        )

    def test_decision_requires_actor_reason_and_known_id(self):
        candidate = self.example()["candidates"][0]["id"]
        for identifier, action, actor, reason in [
            (candidate, "accept", "", "Enough reason"),
            (candidate, "accept", "A", "short"),
            ("unknown", "accept", "A", "Enough reason"),
            (candidate, "post", "A", "Enough reason"),
        ]:
            with self.subTest(action=action), self.assertRaises(ValueError):
                self.store.decide(identifier, action, actor, reason)

    def test_restart_preserves_balances_and_audit(self):
        candidate = next(
            item for item in self.example()["candidates"] if item["method"] == "reference_batch"
        )
        self.store.decide(
            candidate["id"], "accept", "Reviewer", "Verified batch receipt references"
        )
        expected = self.store.report()
        with Store(self.filename) as reopened:
            self.assertEqual(reopened.report(), expected)
        for total in expected["summary"]:
            self.assertEqual(total["accepted_bank"], total["accepted_ledger"])

    def test_parallel_reviewers_cannot_allocate_twice(self):
        choices = [item for item in self.example()["candidates"] if item["ambiguous"]]
        barrier = threading.Barrier(2)
        outcomes = []

        def review(candidate):
            with Store(self.filename) as other:
                barrier.wait(timeout=5)
                try:
                    other.decide(candidate["id"], "accept", "Reviewer", "Concurrency test evidence")
                    outcomes.append("accepted")
                except ValueError:
                    outcomes.append("conflict")

        threads = [threading.Thread(target=review, args=(item,)) for item in choices]
        for thread in threads:
            thread.start()
        for thread in threads:
            thread.join(timeout=10)
            self.assertFalse(thread.is_alive())
        self.assertEqual(sorted(outcomes), ["accepted", "conflict"])
        self.store.verify_audit()

    def test_mutation_trigger_and_tamper_detection(self):
        self.example()
        with self.assertRaises(sqlite3.IntegrityError):
            self.store.db.execute("UPDATE audit SET payload='{}' WHERE sequence=1")
        self.store.db.execute("DROP TRIGGER audit_no_update")
        self.store.db.execute("UPDATE audit SET payload='{}' WHERE sequence=1")
        with self.assertRaisesRegex(ValueError, "verification failed"):
            self.store.verify_audit()
        with self.assertRaises(ValueError):
            self.store.reconcile()

    def test_corrupted_candidate_is_not_accepted(self):
        candidate = self.example()["candidates"][0]
        altered = {key: value for key, value in candidate.items() if key != "ambiguous"}
        altered["cents"] += 1
        self.store.db.execute("DROP TRIGGER candidates_no_update")
        self.store.db.execute(
            "UPDATE candidates SET payload=? WHERE id=?", (json.dumps(altered), candidate["id"])
        )
        with self.assertRaisesRegex(ValueError, "integrity"):
            self.store.decide(candidate["id"], "accept", "Reviewer", "Source evidence checked")

    def test_csv_formula_injection_is_neutralized(self):
        self.ingest("bank", [("=SUM(1)", "10", "+HYPERLINK(1)")])
        handle = StringIO()
        export_csv(self.store.report(), handle)
        rows = list(csv.reader(StringIO(handle.getvalue())))
        self.assertEqual(rows[1][1], "'=SUM(1)")
        self.assertEqual(rows[1][6], "'+HYPERLINK(1)")

    def test_no_implicit_overwrite_of_export(self):
        destination = Path(self.temp.name) / "report.json"
        destination.write_text("preserve me")
        with patch("sys.stderr", new=StringIO()):
            code = main(["--db", self.filename, "export", str(destination)])
        self.assertEqual(code, 2)
        self.assertEqual(destination.read_text(), "preserve me")

    def test_unrelated_sqlite_database_is_not_modified(self):
        path = Path(self.temp.name) / "unrelated.sqlite3"
        with sqlite3.connect(path) as connection:
            connection.execute("CREATE TABLE important(value TEXT)")
        with self.assertRaisesRegex(ValueError, "not an empty"):
            Store(path)


if __name__ == "__main__":
    unittest.main()
