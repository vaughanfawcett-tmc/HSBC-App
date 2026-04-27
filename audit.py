"""Append-only audit log.

Every successful send writes one row. Rows are never rewritten or deleted.
The log proves what was sent and when, with integrity hashes for the source
CSV and the output xlsx.
"""
from __future__ import annotations

import csv
import datetime as dt
import hashlib
from pathlib import Path

FIELDS = [
    "timestamp_utc",
    "user_email",
    "source_csv_sha256",
    "output_xlsx_sha256",
    "month_label",
    "services_total_within",
    "services_total_outside",
    "callbacks_total_within",
    "callbacks_total_outside",
    "unmapped_count",
    "recipient",
    "graph_message_id",
]


def sha256_file(path: str | Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def append_entry(log_path: str | Path, entry: dict) -> None:
    log_path = Path(log_path)
    exists = log_path.exists()
    row = {k: entry.get(k, "") for k in FIELDS}
    if "timestamp_utc" not in entry:
        row["timestamp_utc"] = dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with open(log_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDS)
        if not exists:
            writer.writeheader()
        writer.writerow(row)
