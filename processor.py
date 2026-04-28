"""Pure CSV -> tallies logic. No file I/O, no network, no timestamps.

All functions are deterministic and idempotent: same inputs always produce
identical outputs. This is the contract the tests pin down.
"""
from __future__ import annotations

import re
import warnings
from dataclasses import dataclass, field
from typing import Any

import pandas as pd
import yaml

# pandas warns when str.contains patterns contain capturing groups, but we
# only use them as boolean filters — the warning is noise for this tool.
warnings.filterwarnings(
    "ignore",
    message=r"This pattern is interpreted as a regular expression.*",
    category=UserWarning,
)

UNMAPPED = "<UNMAPPED>"


@dataclass(frozen=True)
class MatchRule:
    field: str
    pattern: re.Pattern


@dataclass(frozen=True)
class ClientRule:
    name: str
    rules: tuple[MatchRule, ...]
    # Lower-cased email domains. A row matches if its Contact ID's @-suffix
    # matches any of these (e.g. "honeywell.com" matches "x@honeywell.com" AND
    # "x@uk.honeywell.com"). Cheaper and far less brittle than regex name matches.
    domains: tuple[str, ...] = ()


@dataclass
class SummaryReport:
    month_label: str
    services: dict[str, tuple[int, int]] = field(default_factory=dict)
    callbacks: dict[str, tuple[int, int]] = field(default_factory=dict)
    unmapped_senders: dict[str, int] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    total_rows: int = 0

    @property
    def services_totals(self) -> tuple[int, int]:
        return _sum_pairs(self.services.values())

    @property
    def callbacks_totals(self) -> tuple[int, int]:
        return _sum_pairs(self.callbacks.values())

    @property
    def unmapped_total(self) -> int:
        return sum(self.unmapped_senders.values())


def _sum_pairs(pairs) -> tuple[int, int]:
    within = outside = 0
    for w, o in pairs:
        within += w
        outside += o
    return within, outside


# ---------- config loading ----------

def load_clients_config(path: str) -> list[ClientRule]:
    with open(path, encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    clients = data.get("clients", [])
    out: list[ClientRule] = []
    for entry in clients:
        name = entry["name"]
        raw_rules = entry.get("rules", [])
        rules = tuple(
            MatchRule(
                field=r["field"],
                pattern=re.compile(r["match"], re.IGNORECASE | re.DOTALL),
            )
            for r in raw_rules
        )
        domains = tuple(
            d.strip().lower().lstrip("@")
            for d in (entry.get("domains") or [])
            if isinstance(d, str) and d.strip()
        )
        out.append(ClientRule(name=name, rules=rules, domains=domains))
    return out


def load_settings(path: str) -> dict[str, Any]:
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


# ---------- transformations ----------

def _email_domain_series(df: pd.DataFrame) -> pd.Series:
    """Lower-cased email domain extracted from Contact ID (the part after @).
    Empty string when Contact ID is missing or has no @."""
    if "Contact ID" not in df.columns:
        return pd.Series([""] * len(df), index=df.index, dtype=object)
    contact = df["Contact ID"].fillna("").astype(str).str.strip().str.lower()
    return contact.str.extract(r"@([^@\s>]+)", expand=False).fillna("")


def assign_client(df: pd.DataFrame, clients: list[ClientRule]) -> pd.DataFrame:
    """Add a 'client' column to df. First matching client wins.

    Match priority for each client (cheapest signal first, most specific last):
      1. Email domain match — Contact ID's @-suffix matches `domains` exactly,
         OR is a sub-domain of one of them (e.g. uk.honeywell.com matches
         honeywell.com).
      2. Regex rule match — any rule's field/pattern matches.

    Unmatched rows get client = UNMAPPED.
    """
    out = df.copy()
    assigned = pd.Series([UNMAPPED] * len(out), index=out.index, dtype=object)
    domain_series = _email_domain_series(out)

    for client in clients:
        if not client.rules and not client.domains:
            continue
        still_unassigned = assigned == UNMAPPED
        if not still_unassigned.any():
            break

        mask = pd.Series([False] * len(out), index=out.index)

        # 1) Domain match (exact OR sub-domain).
        for d in client.domains:
            d_clean = d.lower().lstrip("@").strip()
            if not d_clean:
                continue
            mask = mask | (domain_series == d_clean) | domain_series.str.endswith("." + d_clean)

        # 2) Regex rule match.
        for rule in client.rules:
            if rule.field not in out.columns:
                continue
            col = out[rule.field].fillna("").astype(str)
            mask = mask | col.str.contains(rule.pattern, na=False, regex=True)

        mask = mask & still_unassigned
        assigned.loc[mask] = client.name

    out["client"] = assigned
    return out


def filter_services(df: pd.DataFrame, exclude_groups: list[str]) -> pd.DataFrame:
    """Drop rows whose Group is in exclude_groups (case-insensitive)."""
    if "Group" not in df.columns or not exclude_groups:
        return df
    lowered = {g.strip().lower() for g in exclude_groups}
    group_col = df["Group"].fillna("").astype(str).str.strip().str.lower()
    return df[~group_col.isin(lowered)].copy()


def filter_callbacks(df: pd.DataFrame, rule) -> pd.DataFrame:
    """Keep only rows matching any callback rule.

    Accepts:
      - a single {field, match} dict (legacy)
      - a list of {field, match} dicts — a row matches if ANY rule matches.
    """
    if not rule:
        return df.iloc[0:0].copy()
    rules_list = rule if isinstance(rule, list) else [rule]

    mask = pd.Series([False] * len(df), index=df.index)
    matched_any = False
    for r in rules_list:
        if not isinstance(r, dict) or "field" not in r or "match" not in r:
            continue
        field = r["field"]
        if field not in df.columns:
            continue
        pattern = re.compile(r["match"], re.IGNORECASE | re.DOTALL)
        col = df[field].fillna("").astype(str)
        mask = mask | col.str.contains(pattern, na=False, regex=True)
        matched_any = True

    if not matched_any:
        return df.iloc[0:0].copy()
    return df[mask].copy()


def tally(df: pd.DataFrame) -> dict[str, tuple[int, int]]:
    """Return {client: (within_sla, sla_violated)} keyed on assigned client.

    Rows where Resolution status is empty or something other than the two
    known values are ignored — callers should look at summary warnings.
    """
    if df.empty or "client" not in df.columns or "Resolution status" not in df.columns:
        return {}
    status = df["Resolution status"].fillna("").astype(str).str.strip()
    within_mask = status.str.casefold() == "within sla"
    violated_mask = status.str.casefold() == "sla violated"

    result: dict[str, tuple[int, int]] = {}
    for client in df["client"].dropna().unique():
        rows = df["client"] == client
        w = int((rows & within_mask).sum())
        o = int((rows & violated_mask).sum())
        result[client] = (w, o)
    return result


def unmapped_breakdown(df: pd.DataFrame, identifier_field: str = "Full name") -> dict[str, int]:
    """For unmapped rows, count occurrences per identifier so the reviewer
    knows which senders to add to clients.yaml."""
    if df.empty or "client" not in df.columns:
        return {}
    if identifier_field not in df.columns:
        return {}
    unmapped = df[df["client"] == UNMAPPED]
    if unmapped.empty:
        return {}
    counts = (
        unmapped[identifier_field]
        .fillna("(blank)")
        .astype(str)
        .value_counts()
        .to_dict()
    )
    return {k: int(v) for k, v in counts.items()}


def detect_month_label(df: pd.DataFrame, created_col: str = "Created time") -> str:
    """Return the dominant month in the export as 'Month YYYY'.

    Fresh CRM exports use either ISO 'YYYY-MM-DD HH:MM:SS' or UK 'DD/MM/YYYY'
    depending on tenant settings. We try the default parse first (handles ISO
    correctly) and only fall back to dayfirst=True when the default leaves a
    lot of unparseable rows — which happens with UK strings like '13/03/2026'
    that the default reads as month 13 = NaT. Without this fallback, March
    data uploaded as DD/MM gets misread as the wrong month.
    """
    if created_col not in df.columns or df.empty:
        return ""
    raw = df[created_col]
    total = len(raw.dropna())
    default = pd.to_datetime(raw, errors="coerce").dropna()
    parsed = default
    # If the default parse loses more than 30% of rows, it's almost certainly
    # a DD/MM tenant — retry dayfirst=True. Otherwise stick with the default
    # parse (ISO/US dates that dayfirst would silently corrupt).
    if total and len(default) < 0.7 * total:
        uk = pd.to_datetime(raw, errors="coerce", dayfirst=True).dropna()
        if len(uk) > len(default):
            parsed = uk
    if parsed.empty:
        return ""
    most_common = parsed.dt.to_period("M").value_counts().idxmax()
    return most_common.to_timestamp().strftime("%B %Y")


def summarise(
    df_services: pd.DataFrame,
    df_callbacks: pd.DataFrame,
    template_clients: list[str],
    month_label: str,
    identifier_field: str = "Full name",
) -> SummaryReport:
    """Assemble the full report: services tallies, callbacks tallies,
    unmapped senders, and validation warnings.

    The template_clients list is authoritative: clients in that list get a
    (0, 0) entry even when no tickets landed. Clients not in the list raise
    a warning.
    """
    services_tallies = tally(df_services)
    callbacks_tallies = tally(df_callbacks)

    unmapped_svc = unmapped_breakdown(df_services, identifier_field)
    unmapped_cb = unmapped_breakdown(df_callbacks, identifier_field)
    unmapped_combined: dict[str, int] = {}
    for d in (unmapped_svc, unmapped_cb):
        for k, v in d.items():
            unmapped_combined[k] = unmapped_combined.get(k, 0) + v

    services_final: dict[str, tuple[int, int]] = {}
    callbacks_final: dict[str, tuple[int, int]] = {}
    template_set = set(template_clients)

    for client in template_clients:
        services_final[client] = services_tallies.get(client, (0, 0))
        callbacks_final[client] = callbacks_tallies.get(client, (0, 0))

    warnings: list[str] = []
    for client, (w, o) in services_tallies.items():
        if client == UNMAPPED:
            continue
        if client not in template_set:
            warnings.append(
                f"Client '{client}' has services tickets ({w + o}) but is not in the template."
            )
    for client, (w, o) in callbacks_tallies.items():
        if client == UNMAPPED:
            continue
        if client not in template_set:
            warnings.append(
                f"Client '{client}' has callback tickets ({w + o}) but is not in the template."
            )

    status_col = "Resolution status"
    if status_col in df_services.columns:
        blank = df_services[status_col].fillna("").astype(str).str.strip() == ""
        n = int(blank.sum())
        if n:
            warnings.append(f"{n} ticket(s) have blank Resolution status and were not counted.")

    return SummaryReport(
        month_label=month_label,
        services=services_final,
        callbacks=callbacks_final,
        unmapped_senders=unmapped_combined,
        warnings=warnings,
        total_rows=len(df_services) + len(df_callbacks),
    )
