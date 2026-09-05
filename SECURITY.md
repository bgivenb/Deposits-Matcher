# Security and data handling

This is a local portfolio utility, not a hosted financial service or certified accounting control.

- Workspaces, SQLite WAL/SHM sidecars, and exports contain financial data in plaintext. Use a private directory, appropriate OS permissions, disk encryption, backups, and retention. Never commit real statements.
- The application makes no network, upload, telemetry, or model-provider calls. Package installation is separate and may access a package index.
- Reviewer names are locally supplied labels, not authenticated identities. There is no remote-access authorization boundary. Account/currency matching scopes are not access controls.
- Hash-linked audit events and immutability triggers detect ordinary corruption and discourage accidental editing. A sufficiently privileged user can alter the database and rebuild a chain; no independent immutable anchor exists.
- Imports reject invalid schema, duplicate IDs, conflicting content, invalid dates, malformed/non-finite money, fractional cents, controls, and oversized input. SQL values are parameterized.
- CSV exports prefix formula-like text with an apostrophe. JSON exports retain validated original values, source lineage, decisions, and journal entries. Both can contain sensitive financial data.
- No match posts to a bank or ledger. Acceptance records a local reviewer conclusion only. Corrections/reversals are not implemented; this version must not be your system of record.
- Back up a live workspace through SQLite's backup API, or close all connections before copying it. Do not copy only the main database while ignoring a live WAL.

Report concerns using GitHub private vulnerability reporting if available, or the maintainer's profile contact link. Do not attach real customer/account information to public issues.
