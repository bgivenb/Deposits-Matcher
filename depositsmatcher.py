"""Tkinter interface for Deposits Matcher."""

from __future__ import annotations

import queue
import re
import threading
import tkinter as tk
from decimal import Decimal
from tkinter import filedialog, messagebox, ttk

from matching import MAX_DEPOSITS_PER_LIST, MatchResult, find_maximum_match, parse_amounts


BACKGROUND = "#202124"
PANEL = "#303134"
ENTRY = "#111214"
TEXT = "#f5f5f5"
MUTED = "#b6b6b6"
MATCHED = "#57c785"
UNMATCHED = "#ef6a6a"


def currency(value: Decimal) -> str:
    return f"${value:,.2f}"


class ScrollableList(tk.Frame):
    def __init__(self, parent: tk.Widget, title: str) -> None:
        super().__init__(parent, bg=PANEL)
        tk.Label(
            self,
            text=title,
            font=("Arial", 13, "bold"),
            bg=PANEL,
            fg=TEXT,
        ).pack(anchor="w", padx=12, pady=(12, 6))

        canvas = tk.Canvas(self, borderwidth=0, background=PANEL, highlightthickness=0)
        scrollbar = tk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.body = tk.Frame(canvas, background=PANEL)
        self.body.bind(
            "<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.body, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)


class DepositsMatcherApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Deposits Matcher")
        self.root.geometry("960x720")
        self.root.configure(bg=BACKGROUND)
        self.entries_a: list[tuple[tk.Entry, tk.Label]] = []
        self.entries_b: list[tuple[tk.Entry, tk.Label]] = []
        self.amounts_a: list[Decimal] = []
        self.amounts_b: list[Decimal] = []
        self.result: MatchResult | None = None

        header = tk.Frame(root, bg=BACKGROUND)
        header.pack(fill="x", padx=20, pady=(18, 8))
        tk.Label(
            header,
            text="DEPOSITS MATCHER",
            font=("Arial", 22, "bold"),
            bg=BACKGROUND,
            fg=TEXT,
        ).pack(side="left")
        tk.Button(header, text="Help", command=self.show_help).pack(side="right")

        tk.Label(
            root,
            text="Find the largest exact subset total shared by two deposit lists.",
            bg=BACKGROUND,
            fg=MUTED,
        ).pack(anchor="w", padx=20)

        setup = tk.Frame(root, bg=BACKGROUND)
        setup.pack(fill="x", padx=20, pady=12)
        tk.Label(setup, text="List A count", bg=BACKGROUND, fg=TEXT).grid(row=0, column=0)
        self.count_a = tk.Entry(setup, width=6, bg=ENTRY, fg=TEXT, insertbackground=TEXT)
        self.count_a.insert(0, "6")
        self.count_a.grid(row=0, column=1, padx=(8, 22))
        tk.Label(setup, text="List B count", bg=BACKGROUND, fg=TEXT).grid(row=0, column=2)
        self.count_b = tk.Entry(setup, width=6, bg=ENTRY, fg=TEXT, insertbackground=TEXT)
        self.count_b.insert(0, "6")
        self.count_b.grid(row=0, column=3, padx=(8, 22))
        self.generate_button = tk.Button(
            setup, text="Generate fields", command=self.generate_fields
        )
        self.generate_button.grid(row=0, column=4)

        lists = tk.Frame(root, bg=BACKGROUND)
        lists.pack(fill="both", expand=True, padx=20, pady=6)
        self.list_a = ScrollableList(lists, "LIST A")
        self.list_a.pack(side="left", fill="both", expand=True, padx=(0, 6))
        self.list_b = ScrollableList(lists, "LIST B")
        self.list_b.pack(side="left", fill="both", expand=True, padx=(6, 0))

        actions = tk.Frame(root, bg=BACKGROUND)
        actions.pack(fill="x", padx=20, pady=10)
        tk.Button(
            actions, text="Paste List A", command=lambda: self.paste_values("A")
        ).pack(side="left")
        tk.Button(
            actions, text="Paste List B", command=lambda: self.paste_values("B")
        ).pack(side="left", padx=8)
        self.match_button = tk.Button(
            actions, text="Find maximum match", command=self.calculate
        )
        self.match_button.pack(side="left", padx=(16, 8))
        tk.Button(actions, text="Export .xlsx", command=self.export).pack(side="left")
        self.progress = ttk.Progressbar(actions, mode="indeterminate", length=160)
        self.progress.pack(side="right")

        self.result_text = tk.StringVar(
            value="Generate fields, enter amounts, then run the matcher."
        )
        self.worker_events: queue.Queue = queue.Queue()
        self.active_job = 0
        tk.Label(
            root,
            textvariable=self.result_text,
            justify="left",
            anchor="w",
            bg=PANEL,
            fg=TEXT,
            padx=14,
            pady=12,
        ).pack(fill="x", padx=20, pady=(0, 18))
        self.generate_fields()
        self.root.after(100, self._poll_worker)

    def show_help(self) -> None:
        messagebox.showinfo(
            "Deposits Matcher",
            "Enter non-negative currency values with at most two decimal places. "
            "The matcher finds the largest exact subset total shared by List A and List B.\n\n"
            f"The limit is {MAX_DEPOSITS_PER_LIST} entries per list because subset matching is exponential.",
        )

    def _count(self, entry: tk.Entry) -> int:
        try:
            count = int(entry.get())
        except ValueError as exc:
            raise ValueError("Deposit counts must be whole numbers.") from exc
        if not 1 <= count <= MAX_DEPOSITS_PER_LIST:
            raise ValueError(
                f"Deposit counts must be between 1 and {MAX_DEPOSITS_PER_LIST}."
            )
        return count

    def generate_fields(self) -> None:
        try:
            count_a = self._count(self.count_a)
            count_b = self._count(self.count_b)
        except ValueError as exc:
            messagebox.showerror("Invalid count", str(exc))
            return

        self._replace_entries(self.list_a.body, self.entries_a, count_a, "A")
        self._replace_entries(self.list_b.body, self.entries_b, count_b, "B")
        self.result = None
        self.result_text.set("Enter amounts or paste columns copied from a spreadsheet.")

    def _replace_entries(
        self,
        parent: tk.Frame,
        entries: list[tuple[tk.Entry, tk.Label]],
        count: int,
        prefix: str,
    ) -> None:
        for child in parent.winfo_children():
            child.destroy()
        entries.clear()
        for index in range(count):
            row = tk.Frame(parent, bg=PANEL)
            row.pack(fill="x", padx=12, pady=4)
            tk.Label(
                row, text=f"{prefix}{index + 1}", width=4, bg=PANEL, fg=MUTED
            ).pack(side="left")
            amount = tk.Entry(row, bg=ENTRY, fg=TEXT, insertbackground=TEXT)
            amount.pack(side="left", fill="x", expand=True, padx=6)
            status = tk.Label(row, text="", width=10, bg=PANEL, fg=MUTED)
            status.pack(side="left")
            entries.append((amount, status))

    def paste_values(self, list_name: str) -> None:
        entries = self.entries_a if list_name == "A" else self.entries_b
        try:
            values = [
                value.strip()
                for value in re.split(r"[\t\r\n]+", self.root.clipboard_get())
                if value.strip()
            ]
        except tk.TclError:
            messagebox.showerror("Paste failed", "The clipboard does not contain text.")
            return
        if len(values) > len(entries):
            messagebox.showwarning("Paste truncated", "Extra clipboard values were ignored.")
        for index, (entry, _status) in enumerate(entries):
            entry.delete(0, tk.END)
            if index < len(values):
                entry.insert(0, values[index])

    def calculate(self) -> None:
        try:
            self.amounts_a = parse_amounts(entry.get() for entry, _ in self.entries_a)
            self.amounts_b = parse_amounts(entry.get() for entry, _ in self.entries_b)
        except ValueError as exc:
            messagebox.showerror("Invalid amount", str(exc))
            return

        self.match_button.configure(state="disabled")
        self.generate_button.configure(state="disabled")
        self.progress.start(10)
        self.active_job += 1
        job = self.active_job
        threading.Thread(
            target=self._calculate_worker,
            args=(job, tuple(self.amounts_a), tuple(self.amounts_b)),
            daemon=True,
        ).start()

    def _calculate_worker(
        self, job: int, amounts_a: tuple[Decimal, ...], amounts_b: tuple[Decimal, ...]
    ) -> None:
        try:
            result = find_maximum_match(amounts_a, amounts_b)
        except Exception as exc:
            self.worker_events.put((job, "error", exc))
            return
        self.worker_events.put((job, "result", result))

    def _poll_worker(self) -> None:
        try:
            while True:
                job, kind, value = self.worker_events.get_nowait()
                if job != self.active_job:
                    continue
                if kind == "error":
                    self._calculation_failed(value)
                else:
                    self._render_result(value)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_worker)

    def _calculation_failed(self, exc: Exception) -> None:
        self.progress.stop()
        self.match_button.configure(state="normal")
        self.generate_button.configure(state="normal")
        messagebox.showerror("Matching failed", str(exc))

    def _render_result(self, result: MatchResult) -> None:
        self.progress.stop()
        self.match_button.configure(state="normal")
        self.generate_button.configure(state="normal")
        self.result = result
        matched_a = set(result.matched_indices_a)
        matched_b = set(result.matched_indices_b)
        for index, (_entry, status) in enumerate(self.entries_a):
            is_matched = index in matched_a
            status.configure(
                text="Matched" if is_matched else "Unmatched",
                fg=MATCHED if is_matched else UNMATCHED,
            )
        for index, (_entry, status) in enumerate(self.entries_b):
            is_matched = index in matched_b
            status.configure(
                text="Matched" if is_matched else "Unmatched",
                fg=MATCHED if is_matched else UNMATCHED,
            )

        self.result_text.set(
            f"Matched total: {currency(result.matched_total)}\n"
            f"List A unmatched: {currency(result.unmatched_total_a)}\n"
            f"List B unmatched: {currency(result.unmatched_total_b)}\n"
            f"Matched entries: A {tuple(index + 1 for index in result.matched_indices_a)} ↔ "
            f"B {tuple(index + 1 for index in result.matched_indices_b)}"
        )

    def export(self) -> None:
        if self.result is None:
            messagebox.showinfo("Nothing to export", "Run the matcher first.")
            return
        try:
            from openpyxl import Workbook
        except ImportError:
            messagebox.showerror(
                "Missing dependency",
                "Install dependencies with: pip install -r requirements.txt",
            )
            return

        destination = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")]
        )
        if not destination:
            return

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Deposit Match"
        sheet.append(["List", "Entry", "Amount", "Status"])
        matched_a = set(self.result.matched_indices_a)
        matched_b = set(self.result.matched_indices_b)
        for name, amounts, matched in (
            ("A", self.amounts_a, matched_a),
            ("B", self.amounts_b, matched_b),
        ):
            for index, amount in enumerate(amounts):
                sheet.append(
                    [
                        name,
                        index + 1,
                        float(amount),
                        "Matched" if index in matched else "Unmatched",
                    ]
                )
                sheet.cell(row=sheet.max_row, column=3).number_format = "$#,##0.00"
        sheet.append([])
        sheet.append(["Matched total", float(self.result.matched_total)])
        sheet.cell(row=sheet.max_row, column=2).number_format = "$#,##0.00"
        workbook.save(destination)
        messagebox.showinfo("Export complete", f"Saved {destination}")


if __name__ == "__main__":
    app_root = tk.Tk()
    DepositsMatcherApp(app_root)
    app_root.mainloop()
