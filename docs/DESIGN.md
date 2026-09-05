# Design and invariants

## Why not maximize matched dollars?

Equal totals can be coincidental. The original solver optimizes a mathematical objective, not the truth of a financial relationship. The workbench generates explanations for review instead of selecting the combination that makes the largest balance disappear.

Matching needs no LLM: exact arithmetic, explicit references, bounded date windows, and deterministic search are easier to reproduce. Human judgment remains necessary when evidence is ambiguous.

## Data flow

```text
bank / ledger CSV
    → whole-file validation
    → atomic, idempotent import + source receipt + row lineage
    → versioned, bounded candidate generation
    → review queue (including competing candidates)
    → explicit acceptance / rejection
    → transactional allocation + immutable decision + audit event
    → current-state JSON / formula-safe CSV export
```

## Enforced invariants

1. Money is parsed as Decimal, then stored/compared as integer cents. Totals never cross accounts or currencies.
2. Records use SHA-256 of `[source, account, external_id]`, not amount or file position. Identical amounts remain distinct records.
3. A bad row or conflicting record aborts the entire import. Byte-identical reimports are no-ops; overlapping files retain additional lineage without duplicating records.
4. Every proposal balances exactly. Each batch is reference-backed and date-constrained. Input order cannot decide a winner because the engine never selects one.
5. `BEGIN IMMEDIATE` serializes review writes; unique transaction allocations prevent double consumption. Candidate evidence is revalidated before a decision.
6. Repeating an identical decision is idempotent. Changing its action, actor, or reason is rejected. There is no unsupported undo that erases history.
7. State changes and their audit events commit together or roll back together. UI work runs off the Tk event thread and returns through a queue.
8. Exports use a consistent SQLite snapshot, not half-updated balances from concurrent local decisions.

## Bounded search

One-to-one matching uses an amount index. Batch search uses an account/currency/reference index, then enumerates groups of two through four. Oversized pools are skipped with explicit limitations instead of selecting an arbitrary prefix. Global comparison/proposal budgets stop pathological duplicate-heavy inputs.

This is not a general subset-sum optimizer. Fees, many-to-many settlements, reversals, fuzzy descriptions, and partial references require another reviewed policy. No proposal is not proof that no financial relationship exists.

## Audit journal

Each append includes sequence, UTC timestamp, event type, a local actor label, and details. Its digest is `SHA256(previous_hash + canonical_payload)`, with 64 zeroes as the first predecessor. Canonical JSON sorts keys, uses compact separators, and preserves Unicode. SQL triggers reject ordinary updates/deletes to events, decisions, imported records, and candidates. JSON exports include the full event chain and its head.

This detects journal corruption and accidental mutations. It does not authenticate reviewers, encrypt data, prevent an administrator from rebuilding a chain, or prove that a statement is genuine. A production service needs organization-owned identity, authorization, encryption/retention controls, independent checkpoints, and controlled reversals.

## Build versus adopt

SQLite supplies transactions, locking, crash recovery, and constraints; Python supplies Decimal, CSV, hashing, and Tk. Custom code is restricted to reconciliation policy, evidence representation, and the review workflow. No custom database, web framework, agent runtime, or probabilistic matching service is necessary for this local problem.

## What would justify a next version?

- Demand for a bank export format: an explicit adapter with contract fixtures.
- A need to correct accepted decisions: reversal events with allocation-release semantics, not mutable history.
- Remote/shared access: an authenticated service and authorization boundary before exposing data.
- Measured policy failures on real batches: a separately versioned policy with labeled evaluations, not an unbounded search limit.
