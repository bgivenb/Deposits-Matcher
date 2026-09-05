"""Open the native workbench with disposable, explicitly synthetic example data."""

from pathlib import Path
import sys
import tempfile
import tkinter as tk

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from workbench import Workbench

ROOT = Path(__file__).resolve().parents[1]


def main():
    with tempfile.TemporaryDirectory(prefix="deposits-preview-") as directory:
        root = tk.Tk()
        app = Workbench(root)
        root.title("Deposits Matcher — Synthetic demo (temporary workspace)")
        app.filename = str(Path(directory) / "synthetic-demo.sqlite3")

        def load(store):
            for source in ["bank", "ledger"]:
                store.import_statement(ROOT / f"examples/{source}.csv", source)
            store.reconcile()
            return "Synthetic demo — temporary workspace"

        app.task(load)

        def select_example():
            if app.busy:
                root.after(100, select_example)
            elif app.report:
                item = next(
                    candidate
                    for candidate in app.report["candidates"]
                    if candidate["reason"]["reference"] == "BATCH-7"
                )
                app.candidates.selection_set(item["id"])
                app.show_candidate()

        root.after(200, select_example)
        root.mainloop()


if __name__ == "__main__":
    main()
