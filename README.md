# Deposits Matcher

[![Tests](https://github.com/bgivenb/Deposits-Matcher/actions/workflows/test.yml/badge.svg)](https://github.com/bgivenb/Deposits-Matcher/actions/workflows/test.yml)

Deposits Matcher is a small desktop utility for reconciling two lists of currency amounts. It finds the largest exact subset total shared by both lists, marks the entries involved, reports what remains unmatched, and can export the result to Excel.

> This is a hobby utility, not accounting advice or a substitute for review in a system of record.

## Why the matching core is careful

- Currency is parsed with Python `Decimal` and compared as integer cents—never binary floating point.
- Duplicate amounts retain distinct entry identities.
- Invalid, negative, non-finite, and fractional-cent inputs are rejected.
- The algorithm finds the maximum common subset total rather than stopping at the first match.
- Matching logic is independent of Tkinter and covered by unit tests.

## Run from source

Requirements: Python 3.11 or newer. Tkinter is included with the standard Python installers on Windows and macOS.

```bash
python -m venv .venv
source .venv/bin/activate  # Windows PowerShell: .venv\Scripts\Activate.ps1
pip install -r requirements.txt
python depositsmatcher.py
```

## Validate

```bash
python -m unittest discover -s tests -v
```

The subset search is exponential, so the interface intentionally limits each list to 20 entries. This makes the boundary explicit instead of implying that arbitrarily large reconciliations will finish quickly.

## Legacy archives

The two checked-in macOS ZIP files are historical snapshots and are not the canonical source or current release channel. They have not yet been reproduced from this revised source; use the source version above for evaluation. Future binaries should be built reproducibly and published as versioned GitHub Release assets with checksums.

## License

This repository currently uses the [Creative Commons Attribution-NonCommercial 4.0 license](LICENSE). That choice is unusual for software and should be reviewed before wider reuse or release packaging.
