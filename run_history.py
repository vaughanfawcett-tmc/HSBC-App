"""Per-run input/output archive and rejection feedback log.

A "run" is one upload-to-decision cycle. We keep:
  - the source CSV exactly as Bethany uploaded it
  - the generated xlsx (which may or may not have been sent)
  - a row in runs.csv recording what happened
  - if rejected: a row in feedback.csv with the reason text

Files live under {root}/runs/{YYYYMMDD-HHMMSS}-{short_sha}/ so each run is
self-contained for audit and re-investigation.
"""
from __future__ import annotations

import csv
import datetime as dt
import os
from pathlib import Path

RUN_FIELDS = [
    "run_id",
    "timestamp_utc",
    "outcome",            # "sent" | "rejected" | "draft"
    "month_label",
    "source_filename",
    "source_csv_sha256",
    "output_xlsx_sha256",
    "services_total_within",
    "services_total_outside",
    "callbacks_total_within",
    "callbacks_total_outside",
    "unmapped_count",
    "recipient",
    "graph_message_id",
    "user_email",
    "rejection_reason",
    "source_path",
    "output_path",
]

FEEDBACK_FIELDS = [
    "timestamp_utc",
    "run_id",
    "user_email",
    "month_label",
    "reason",
]


def _root() -> Path:
    """Where to keep the archive. Render mounts a persistent disk at /var/data;
    locally we fall back to ./runs."""
    explicit = os.environ.get("HSBC_SLA_RUNS_DIR")
    if explicit:
        return Path(explicit)
    if Path("/var/data").exists() and os.access("/var/data", os.W_OK):
        return Path("/var/data/runs")
    return Path(__file__).parent / "runs"


def make_run_id(source_sha: str | None = None) -> str:
    stamp = dt.datetime.now(dt.timezone.utc).strftime("%Y%m%d-%H%M%S")
    short = (source_sha or "").replace("-", "")[:8] or "anon"
    return f"{stamp}-{short}"


def archive_run(
    run_id: str,
    source_filename: str,
    source_bytes: bytes,
    output_filename: str,
    output_bytes: bytes,
) -> tuple[Path, Path]:
    """Persist the upload and the generated xlsx for this run.
    Returns (source_path, output_path)."""
    run_dir = _root() / run_id
    run_dir.mkdir(parents=True, exist_ok=True)
    src_path = run_dir / source_filename
    out_path = run_dir / output_filename
    src_path.write_bytes(source_bytes)
    out_path.write_bytes(output_bytes)
    return src_path, out_path


def append_run_row(entry: dict) -> None:
    """Append one row to runs.csv (the index of every run)."""
    log_path = _root() / "runs.csv"
    log_path.parent.mkdir(parents=True, exist_ok=True)
    exists = log_path.exists()
    row = {k: entry.get(k, "") for k in RUN_FIELDS}
    if "timestamp_utc" not in entry:
        row["timestamp_utc"] = dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")
    with open(log_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=RUN_FIELDS)
        if not exists:
            writer.writeheader()
        writer.writerow(row)


def append_feedback(run_id: str, user_email: str, month_label: str, reason: str) -> None:
    """Append a rejection reason to feedback.csv."""
    log_path = _root() / "feedback.csv"
    log_path.parent.mkdir(parents=True, exist_ok=True)
    exists = log_path.exists()
    row = {
        "timestamp_utc": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        "run_id": run_id,
        "user_email": user_email or "",
        "month_label": month_label or "",
        "reason": (reason or "").strip(),
    }
    with open(log_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=FEEDBACK_FIELDS)
        if not exists:
            writer.writeheader()
        writer.writerow(row)
