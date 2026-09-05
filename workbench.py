"""Native reconciliation review workbench. No web server, telemetry, or network access."""

import json
from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from reconcile.__main__ import export_csv
from reconcile.domain import money
from reconcile.store import Store


class Workbench:
    def __init__(self, root):
        self.root = root
        root.title("Deposits Matcher — Reconciliation Workbench")
        root.geometry("1180x780")
        root.minsize(850, 600)
        self.filename = None
        self.report = None
        self.events = queue.Queue()
        self.busy = False
        root.protocol("WM_DELETE_WINDOW", self.close)
        self.buttons = []
        header = ttk.Frame(root, padding=12)
        header.pack(fill="x")
        ttk.Label(header, text="Deposits Matcher", font=("Arial", 18, "bold")).pack(side="left")
        ttk.Label(header, text="Local reconciliation • Human-reviewed decisions").pack(side="right")
        toolbar = ttk.Frame(root, padding=(12, 0, 12, 8))
        toolbar.pack(fill="x")
        for title, command in [
            ("Open / create workspace", self.open_workspace),
            ("Import bank CSV", lambda: self.import_csv("bank")),
            ("Import ledger CSV", lambda: self.import_csv("ledger")),
            ("Find candidates", self.match),
            ("Export report", self.export),
        ]:
            button = ttk.Button(toolbar, text=title, command=command)
            button.pack(side="left", padx=(0, 6))
            self.buttons.append(button)
        self.status = tk.StringVar(
            value="Open a local SQLite workspace to begin. No data leaves this computer."
        )
        ttk.Label(root, textvariable=self.status, padding=(12, 4), wraplength=1100).pack(fill="x")
        self.summary = tk.StringVar(
            value="Suggestions are not approvals. This tool does not post to a bank or ledger."
        )
        ttk.Label(root, textvariable=self.summary, padding=(12, 8), wraplength=1100).pack(fill="x")
        notebook = ttk.Notebook(root)
        notebook.pack(fill="both", expand=True, padx=12, pady=8)
        review = ttk.Frame(notebook, padding=8)
        notebook.add(review, text="Review queue")
        self.candidates = self.table(
            review, ("status", "method", "account", "amount", "bank", "ledger", "ambiguity")
        )
        self.candidates.bind("<<TreeviewSelect>>", self.show_candidate)
        details_frame = ttk.Frame(review)
        details_frame.pack(fill="x", pady=8)
        self.details = tk.Text(
            details_frame, height=14, wrap="word", font=("Courier", 11), state="disabled"
        )
        details_scroll = ttk.Scrollbar(details_frame, orient="vertical", command=self.details.yview)
        self.details.configure(yscrollcommand=details_scroll.set)
        details_scroll.pack(side="right", fill="y")
        self.details.pack(side="left", fill="x", expand=True)
        actions = ttk.Frame(review)
        actions.pack(fill="x")
        for title, action in [("Accept selected…", "accept"), ("Reject selected…", "reject")]:
            button = ttk.Button(
                actions, text=title, command=lambda value=action: self.decide(value)
            )
            button.pack(side="left", padx=(0, 8))
            self.buttons.append(button)
        transactions = ttk.Frame(notebook, padding=8)
        notebook.add(transactions, text="Transactions")
        self.transactions = self.table(
            transactions,
            ("source", "id", "account", "date", "amount", "reference", "status", "candidates"),
        )
        evidence = ttk.Frame(notebook, padding=8)
        notebook.add(evidence, text="Import & audit evidence")
        self.evidence = tk.Text(evidence, wrap="word", font=("Courier", 11), state="disabled")
        self.evidence.pack(fill="both", expand=True)
        root.after(100, self.poll)

    def close(self):
        if self.busy:
            messagebox.showinfo(
                "Operation in progress",
                "Wait for the local operation to finish before closing the workspace.",
            )
            return
        self.root.destroy()

    @staticmethod
    def table(parent, columns):
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")
        for column in columns:
            tree.heading(column, text=column.replace("_", " ").title())
            tree.column(column, width=125, minwidth=70)
        scroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        tree.pack(side="left", fill="both", expand=True)
        return tree

    @staticmethod
    def write_text(widget, text):
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        widget.configure(state="disabled")

    def task(self, operation):
        if self.busy:
            return
        if not self.filename:
            messagebox.showinfo("Workspace required", "Open or create a workspace first.")
            return
        self.busy = True
        for button in self.buttons:
            button.configure(state="disabled")
        filename = self.filename
        self.status.set("Working locally…")

        def worker():
            try:
                with Store(filename) as store:
                    outcome = operation(store)
                    report = store.report()
                self.events.put((report, outcome, None))
            except Exception as exc:
                self.events.put((None, None, str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def poll(self):
        try:
            report, outcome, error = self.events.get_nowait()
        except queue.Empty:
            pass
        else:
            self.busy = False
            for button in self.buttons:
                button.configure(state="normal")
            if error:
                self.status.set(f"Operation failed: {error}")
                messagebox.showerror("Operation failed", error)
            else:
                self.report = report
                self.refresh()
                self.status.set(
                    f"{Path(self.filename).name} • {outcome or 'Workspace loaded'} • Audit chain verified ({report['audit']['events']} events)"
                )
        self.root.after(100, self.poll)

    def open_workspace(self):
        filename = filedialog.asksaveasfilename(
            title="Select an existing workspace or name a new one",
            defaultextension=".sqlite3",
            filetypes=[("Reconciliation workspace", "*.sqlite3")],
            confirmoverwrite=False,
        )
        if filename:
            self.filename = filename
            self.report = None
            self.summary.set("Loading workspace…")
            self.candidates.delete(*self.candidates.get_children())
            self.transactions.delete(*self.transactions.get_children())
            self.write_text(self.details, "")
            self.write_text(self.evidence, "")
            self.task(lambda store: "Workspace loaded")

    def import_csv(self, source):
        if not self.filename:
            messagebox.showinfo("Workspace required", "Open or create a workspace first.")
            return
        filename = filedialog.askopenfilename(
            title=f"Import {source} CSV", filetypes=[("UTF-8 CSV", "*.csv")]
        )
        if filename:

            def operation(store):
                result = store.import_statement(filename, source)
                return f"Imported {result['inserted']} new records; {result['existing']} already present"

            self.task(operation)

    def match(self):
        days = simpledialog.askinteger(
            "Posting-date window",
            "Maximum posting-date difference (days):",
            initialvalue=3,
            minvalue=0,
            maxvalue=31,
        )
        if days is not None:

            def operation(store):
                result = store.reconcile(days)
                return f"Run {result['run_id']}: {len(result['candidates'])} suggestions; {len(result['limits'])} search limitations"

            self.task(operation)

    def refresh(self):
        report = self.report
        index = {row["uid"]: row for row in report["transactions"]}
        self.candidates.delete(*self.candidates.get_children())
        for item in sorted(
            report["candidates"], key=lambda item: (item["status"] != "pending", item["id"])
        ):
            self.candidates.insert(
                "",
                "end",
                iid=item["id"],
                values=(
                    item["status"],
                    item["method"].replace("_", " "),
                    item["account"],
                    f"{money(item['cents'])} {item['currency']}",
                    ", ".join(index[uid]["external_id"] for uid in item["bank"]),
                    ", ".join(index[uid]["external_id"] for uid in item["ledger"]),
                    "Competing candidates" if item["ambiguous"] else "",
                ),
            )
        self.transactions.delete(*self.transactions.get_children())
        for row in report["transactions"]:
            self.transactions.insert(
                "",
                "end",
                values=(
                    row["source"],
                    row["external_id"],
                    row["account"],
                    row["date"],
                    f"{money(row['cents'])} {row['currency']}",
                    row["reference"],
                    row["review_state"],
                    row["pending_candidates"],
                ),
            )
        self.write_text(
            self.details,
            "Select a candidate to inspect its evidence. Every acceptance requires a reviewer and reason.",
        )
        self.write_text(
            self.evidence,
            json.dumps(
                {
                    "imports": report["imports"],
                    "audit": report["audit"],
                    "policy": report["run"]["policy"] if report["run"] else None,
                    "search_limits": report["run"]["limits"] if report["run"] else [],
                    "summary": report["summary"],
                },
                indent=2,
            ),
        )
        unresolved = sum(row["status"] == "unresolved" for row in report["transactions"])
        pending = sum(item["status"] == "pending" for item in report["candidates"])
        limits = len(report["run"]["limits"]) if report["run"] else 0
        self.summary.set(
            f"{len(report['transactions'])} transactions • {unresolved} unresolved • {pending} pending suggestions • {limits} search limitations. Amounts are never combined across currencies or accounts."
        )

    def selected(self):
        ids = self.candidates.selection()
        return next(
            (
                item
                for item in (self.report or {}).get("candidates", [])
                if ids and item["id"] == ids[0]
            ),
            None,
        )

    def show_candidate(self, _event=None):
        item = self.selected()
        if item:
            members = set(item["bank"] + item["ledger"])
            lines = [
                f"{item['method'].replace('_', ' ')} | {money(item['cents'])} {item['currency']} | {item['status']}",
                f"Candidate: {item['id']}",
                f"Posting window: {item['reason']['date_window_days']} days. Reference: {item['reason']['reference'] or '(amount/date only — identify the source before accepting)'}",
            ]
            if item["ambiguous"]:
                lines.append("CAUTION: one or more records also appear in competing candidates.")
            imports = {row["id"]: row for row in self.report["imports"]}
            for row in sorted(
                self.report["transactions"],
                key=lambda record: (record["source"], record["external_id"]),
            ):
                if row["uid"] not in members:
                    continue
                lines.append(
                    f"\n{row['source'].upper()} {row['external_id']} | {row['date']} | {money(row['cents'])} {row['currency']} | {row['reference'] or 'No reference'}\n  {row['description']}"
                )
                for origin in self.report["lineage"]:
                    if origin["uid"] == row["uid"]:
                        lines.append(
                            f"  Source: {imports[origin['import_id']]['filename']}, CSV record row {origin['row_number']}"
                        )
            if item["decision"]:
                lines.append(
                    f"\nReviewer: {item['decision']['actor']}\nReason: {item['decision']['reason']}"
                )
            self.write_text(self.details, "\n".join(lines))

    def decide(self, action):
        item = self.selected()
        if not item or item["status"] != "pending":
            messagebox.showinfo(
                "Select a pending candidate", "Choose an undecided candidate in the review queue."
            )
            return
        if not messagebox.askyesno(
            "Confirm immutable decision",
            f"{action.title()} {money(item['cents'])} {item['currency']}?\n\nThis records a local review decision only, not a bank or ledger posting. Decisions cannot be edited or undone in this version.",
        ):
            return
        actor = simpledialog.askstring(
            "Reviewer", "Reviewer name (local attribution, not authenticated identity):"
        )
        if not actor:
            return
        reason = simpledialog.askstring(
            "Decision reason", "Why is this decision justified? At least 8 characters:"
        )
        if reason:
            self.task(
                lambda store: store.decide(item["id"], action, actor, reason)["action"]
                + " recorded"
            )

    def export(self):
        if not self.report:
            return
        filename = filedialog.asksaveasfilename(
            title="Export a new report file",
            defaultextension=".json",
            filetypes=[("Audit report JSON", "*.json"), ("Transaction CSV", "*.csv")],
        )
        if not filename:
            return
        try:
            # Refresh before exporting so another local reviewer cannot leave this view stale.
            with Store(self.filename) as store:
                report = store.report()
            with open(filename, "x", encoding="utf-8", newline="") as handle:
                if filename.lower().endswith(".csv"):
                    export_csv(report, handle)
                else:
                    json.dump(report, handle, indent=2)
                    handle.write("\n")
            messagebox.showinfo("Exported", "Saved report. Treat it as sensitive financial data.")
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))


if __name__ == "__main__":
    root = tk.Tk()
    Workbench(root)
    root.mainloop()
