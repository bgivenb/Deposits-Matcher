"""Run with python -I after installation to prove the installed package works."""

from pathlib import Path
import tempfile

import reconcile
from reconcile.store import Store

ROOT = Path(__file__).resolve().parents[1]


def main():
    assert not Path(reconcile.__file__).resolve().is_relative_to(ROOT), (
        "Loaded source checkout instead of installed package"
    )
    with tempfile.TemporaryDirectory() as directory:
        filename = Path(directory) / "installed.sqlite3"
        with Store(filename) as store:
            for source in ["bank", "ledger"]:
                store.import_statement(ROOT / f"examples/{source}.csv", source)
            report = store.reconcile()
            assert len(report["candidates"]) == 6
            candidate = next(
                item for item in report["candidates"] if item["reason"]["reference"] == "INV-100"
            )
            store.decide(
                candidate["id"],
                "accept",
                "installation-test",
                "Synthetic installation verification",
            )
        with Store(filename) as store:
            report = store.report()
            assert sum(row["status"] == "accepted" for row in report["transactions"]) == 2
            assert report["audit"]["events"] == 4
    print(
        f"Verified installed Deposits Matcher {reconcile.__version__}: import, propose, review, reopen, audit"
    )


if __name__ == "__main__":
    main()
