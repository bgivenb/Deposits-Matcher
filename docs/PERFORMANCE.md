# Performance evidence

Run `python scripts/benchmark.py` to print the current interpreter, platform, elapsed time, and peak traced Python allocations.

Measured September 4, 2026, Python 3.11.15, macOS 26.5.1, Apple Silicon:

| Workload | Result |
| --- | --- |
| Synthetic input | 5,000 bank + 5,000 ledger records |
| Distribution | Unique exact amounts/references; one account, USD |
| Proposals | 5,000 |
| Candidate comparisons | 15,000 |
| Elapsed engine time | 0.452 seconds |
| Peak traced Python allocations | 16.62 MiB |

The run used `tracemalloc`, which affects timing; memory is not total process RSS. This measures candidate generation, not CSV import, database writes, desktop rendering, or human review. It is one local measurement, not a percentile distribution, SLA, or production-scale claim. Subsequent runs vary.

Duplicate-heavy inputs are different workloads. The engine bounds comparisons and group enumeration and reports partial/skipped searches. Tests exercise these limits rather than extrapolating this favorable workload into a general scalability claim.
