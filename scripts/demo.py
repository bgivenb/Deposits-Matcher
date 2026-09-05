"""Exercise the synthetic workflow, including a clearly simulated review decision."""

import json
from pathlib import Path
import sys
import tempfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from reconcile.store import Store

ROOT = Path(__file__).resolve().parents[1]


def main():
    with tempfile.TemporaryDirectory(prefix="deposits-demo-") as directory:
        filename = Path(directory) / "demo.sqlite3"
        with Store(filename) as store:
            for source in ["bank", "ledger"]:
                store.import_statement(ROOT / f"examples/{source}.csv", source)
            result = store.reconcile()
            assert len(result["candidates"]) == 6
            assert sum(item["ambiguous"] for item in result["candidates"]) == 2
            item = next(
                item for item in result["candidates"] if item["reason"]["reference"] == "INV-100"
            )
            store.decide(
                item["id"],
                "accept",
                "synthetic-demo-reviewer",
                "Simulated review of the INV-100 fixture; not a real financial decision",
            )
        with Store(filename) as reopened:
            report = reopened.report()
            assert sum(row["status"] == "accepted" for row in report["transactions"]) == 2
            print(
                json.dumps(
                    {
                        "dataset": "synthetic",
                        "transactions": len(report["transactions"]),
                        "proposals": len(report["candidates"]),
                        "ambiguous_proposals": sum(
                            item["ambiguous"] for item in report["candidates"]
                        ),
                        "simulated_accepted_matches": 1,
                        "audit_events": report["audit"]["events"],
                        "restart_verified": True,
                        "summary": report["summary"],
                    },
                    indent=2,
                )
            )


if __name__ == "__main__":
    main()
