# Changelog

## 3.0.0

- Add a native reconciliation workbench and dependency-free CLI.
- Import bank/ledger CSVs with stable IDs, atomic validation, content receipts, and row lineage.
- Propose reference/date/amount matches and bounded one-to-many/many-to-one batches.
- Expose competing candidates, unevaluated records, and search limits without automatic approvals.
- Persist immutable decisions and conflict-safe allocations in SQLite with a hash-linked audit journal.
- Export snapshot-consistent JSON evidence and formula-safe CSV without overwriting files.
- Add synthetic scenarios, race/restart/corruption tests, real Tk tests, and a reproducible benchmark.
- Tighten the original calculator's amount parser; retain its separate bounded mathematical mode and historical binaries.

## Earlier source revision

- Separate integer-cent maximum-subset matching from the original Tkinter UI and add unit tests.
