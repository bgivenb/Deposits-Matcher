"""Real Tk widget checks, opt-in because they require a display server."""

import gc
import os
from pathlib import Path
import tempfile
import time
import unittest
from unittest.mock import patch


@unittest.skipUnless(
    os.environ.get("DEPOSITS_GUI_TEST") == "1",
    "Set DEPOSITS_GUI_TEST=1 with a display to exercise Tk",
)
class WorkbenchTests(unittest.TestCase):
    def setUp(self):
        import tkinter as tk
        from workbench import Workbench

        self.temp = tempfile.TemporaryDirectory()
        self.root = tk.Tk()
        self.root.withdraw()
        self.app = Workbench(self.root)
        self.app.filename = str(Path(self.temp.name) / "workspace.sqlite3")
        self.addCleanup(self.temp.cleanup)
        self.addCleanup(self.close_widgets)

    def close_widgets(self):
        self.root.destroy()
        self.app = None
        self.root = None
        # Multiple Tk interpreters are created in this suite. Collect their
        # widget cycles on the owning thread, not in a later SQLite worker.
        gc.collect()

    def wait_for_worker(self):
        deadline = time.monotonic() + 5
        while self.app.busy and time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.01)
        self.assertFalse(self.app.busy)
        self.assertIsNone(self.app.worker)

    def load_example(self):
        root = Path(__file__).resolve().parents[1]

        def operation(store):
            for source in ["bank", "ledger"]:
                store.import_statement(root / f"examples/{source}.csv", source)
            store.reconcile()
            return "Synthetic fixture loaded"

        self.app.task(operation)
        self.wait_for_worker()

    def test_review_queue_and_traceable_details(self):
        self.load_example()
        self.assertEqual(len(self.app.candidates.get_children()), 6)
        self.assertEqual(len(self.app.transactions.get_children()), 18)
        candidate = next(
            item
            for item in self.app.report["candidates"]
            if item["reason"]["reference"] == "INV-100"
        )
        self.app.candidates.selection_set(candidate["id"])
        self.app.show_candidate()
        details = self.app.details.get("1.0", "end")
        self.assertIn("BANK B-001", details)
        self.assertIn("LEDGER L-001", details)
        self.assertIn("bank.csv, CSV record row 2", details)

    def test_review_dialog_records_decision_and_refreshes(self):
        self.load_example()
        candidate = next(
            item
            for item in self.app.report["candidates"]
            if item["reason"]["reference"] == "INV-100"
        )
        self.app.candidates.selection_set(candidate["id"])
        with (
            patch("workbench.messagebox.askyesno", return_value=True),
            patch(
                "workbench.simpledialog.askstring",
                side_effect=["Test reviewer", "Confirmed source evidence"],
            ),
        ):
            self.app.decide("accept")
        self.wait_for_worker()
        self.assertEqual(
            sum(item["status"] == "accepted" for item in self.app.report["candidates"]), 1
        )

    def test_failure_reenables_controls(self):
        def failure(store):
            raise ValueError("Injected failure")

        with patch("workbench.messagebox.showerror") as dialog:
            self.app.task(failure)
            self.wait_for_worker()
            dialog.assert_called_once()
        self.assertTrue(all(str(button["state"]) == "normal" for button in self.app.buttons))

    def test_close_waits_for_a_running_operation(self):
        self.app.busy = True
        with patch("workbench.messagebox.showinfo") as dialog:
            self.app.close()
        dialog.assert_called_once()
        self.assertTrue(self.root.winfo_exists())


@unittest.skipUnless(os.environ.get("DEPOSITS_GUI_TEST") == "1", "Requires a display")
class LegacyCalculatorTests(unittest.TestCase):
    def test_inputs_are_locked_during_calculation_and_edits_invalidate_export(self):
        import tkinter as tk
        from depositsmatcher import DepositsMatcherApp

        root = tk.Tk()
        root.withdraw()

        def cleanup():
            root.destroy()
            gc.collect()

        self.addCleanup(cleanup)
        app = DepositsMatcherApp(root)
        for entries in [app.entries_a, app.entries_b]:
            for index, (entry, _) in enumerate(entries):
                entry.insert(0, "10" if index == 0 else "0")
        app.calculate()
        self.assertTrue(
            all(str(entry["state"]) == "disabled" for entry, _ in app.entries_a + app.entries_b)
        )
        deadline = time.monotonic() + 5
        while app.result is None and time.monotonic() < deadline:
            root.update()
            time.sleep(0.01)
        self.assertIsNotNone(app.result)
        self.assertTrue(
            all(str(entry["state"]) == "normal" for entry, _ in app.entries_a + app.entries_b)
        )
        app.invalidate_result()
        self.assertIsNone(app.result)
        self.assertTrue(
            all(str(status["text"]) == "" for _, status in app.entries_a + app.entries_b)
        )


if __name__ == "__main__":
    unittest.main()
