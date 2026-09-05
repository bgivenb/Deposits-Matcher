"""Command-line reconciliation; no network services or model credentials required."""

import argparse
import csv
import json
import sqlite3
import sys

from .domain import money
from .store import Store


def safe_cell(value):
    text = str(value)
    return (
        "'" + text
        if text.lstrip().startswith(("=", "+", "-", "@")) or text.startswith(("\t", "\r", "\n"))
        else text
    )


def export_csv(report, handle):
    writer = csv.writer(handle)
    writer.writerow(
        [
            "source",
            "id",
            "account",
            "date",
            "amount",
            "currency",
            "reference",
            "description",
            "status",
            "pending_candidates",
            "accepted_candidate",
            "review_state",
        ]
    )
    for row in report["transactions"]:
        writer.writerow(
            [
                safe_cell(value)
                for value in (
                    row["source"],
                    row["external_id"],
                    row["account"],
                    row["date"],
                    money(row["cents"]),
                    row["currency"],
                    row["reference"],
                    row["description"],
                    row["status"],
                    row["pending_candidates"],
                    row["accepted_candidate"] or "",
                    row["review_state"],
                )
            ]
        )


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Deposits Matcher — local, human-reviewed reconciliation"
    )
    parser.add_argument(
        "--db", required=True, help="SQLite workspace file; contains sensitive statement data"
    )
    commands = parser.add_subparsers(dest="command", required=True)
    ingest = commands.add_parser("import", help="Atomically import a UTF-8 CSV")
    ingest.add_argument("source", choices=["bank", "ledger"])
    ingest.add_argument("file")
    match = commands.add_parser("match", help="Propose matches without accepting any")
    match.add_argument("--window-days", type=int, default=3)
    decision = commands.add_parser("decide", help="Record an immutable local reviewer decision")
    decision.add_argument("candidate_id")
    decision.add_argument("action", choices=["accept", "reject"])
    decision.add_argument("--actor", required=True)
    decision.add_argument("--reason", required=True)
    commands.add_parser("report", help="Show the current workspace report as JSON")
    commands.add_parser("verify", help="Verify the append-only audit hash chain")
    export = commands.add_parser("export", help="Write a new report file; refuses existing files")
    export.add_argument("file")
    export.add_argument("--format", choices=["json", "csv"], default="json")
    args = parser.parse_args(argv)
    try:
        with Store(args.db) as store:
            if args.command == "import":
                result = store.import_statement(args.file, args.source)
            elif args.command == "match":
                result = store.reconcile(args.window_days)
            elif args.command == "decide":
                result = store.decide(args.candidate_id, args.action, args.actor, args.reason)
            elif args.command == "verify":
                result = store.verify_audit()
            else:
                result = store.report()
                if args.command == "export":
                    with open(args.file, "x", encoding="utf-8", newline="") as handle:
                        if args.format == "csv":
                            export_csv(result, handle)
                        else:
                            json.dump(result, handle, indent=2, ensure_ascii=False)
                            handle.write("\n")
                    result = {"exported": args.file, "format": args.format}
        print(json.dumps(result, indent=2, ensure_ascii=False))
        return 0
    except (ValueError, OSError, sqlite3.Error) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
