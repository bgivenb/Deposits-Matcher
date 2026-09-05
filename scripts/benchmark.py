"""Reproducible synthetic workload measurement; not a production throughput claim."""

import json
import platform
from pathlib import Path
import sys
import time
import tracemalloc

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from reconcile.domain import Transaction, digest, propose


def main():
    rows = [
        Transaction(
            digest([source, index]),
            source,
            str(index),
            "2026-08-03",
            index + 1,
            "USD",
            "operating",
            str(index),
            "Synthetic indexed reference workload",
        )
        for source in ["bank", "ledger"]
        for index in range(5000)
    ]
    tracemalloc.start()
    started = time.perf_counter()
    report = propose(rows)
    elapsed = time.perf_counter() - started
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    assert len(report["candidates"]) == 5000 and not report["limits"]
    print(
        json.dumps(
            {
                "python": platform.python_version(),
                "platform": platform.platform(),
                "workload": "synthetic unique amount/reference; no ambiguous batch search",
                "transactions": len(rows),
                "candidates": len(report["candidates"]),
                "comparisons": report["comparisons"],
                "elapsed_seconds": round(elapsed, 4),
                "peak_traced_python_mib": round(peak / 1024**2, 2),
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
