"""Pin down processor semantics.

Covers the edge cases that matter for the monthly report:
    - unmapped senders land in <UNMAPPED>
    - first matching rule wins when a row could go to two clients
    - Resolution status is case-insensitive and ignores blanks
    - clients in the template with zero tickets come out as (0, 0), not missing
"""
from pathlib import Path

import pandas as pd
import pytest

import processor

FIXTURE = Path(__file__).parent / "fixtures" / "sample_export.csv"


def _load_fixture() -> pd.DataFrame:
    return pd.read_csv(FIXTURE, dtype=str, keep_default_na=False)


def _clients():
    return processor.load_clients_config(
        str(Path(__file__).parent.parent / "config" / "clients.yaml")
    )


def test_load_clients_config_has_expected_shape():
    clients = _clients()
    names = [c.name for c in clients]
    assert "Honeywell UK" in names
    assert "Honeywell International" in names
    assert "Ashfield Healthcare" in names
    # Each client has at least one rule
    for c in clients:
        assert len(c.rules) >= 1


def test_filter_services_drops_updates():
    df = _load_fixture()
    out = processor.filter_services(df, ["Updates"])
    assert (out["Group"] == "Updates").sum() == 0
    assert len(out) == len(df) - 1  # one Updates row in the fixture


def test_assign_client_unmapped_goes_to_sentinel():
    df = _load_fixture()
    df = processor.filter_services(df, ["Updates"])  # mirror real flow
    df2 = processor.assign_client(df, _clients())
    unmapped = df2[df2["client"] == processor.UNMAPPED]
    # Ticket 1008 and 1010 have unknown senders
    assert set(unmapped["Ticket ID"]) == {"1008", "1010"}


def test_assign_client_first_rule_wins_for_honeywell_split():
    # honeywell.uk@... must land on Honeywell UK (the specific rule comes first)
    df = _load_fixture()
    df2 = processor.assign_client(df, _clients())
    hw_uk = df2[df2["Contact ID"] == "honeywell.uk@fleetlogistics.com"]
    assert not hw_uk.empty
    assert (hw_uk["client"] == "Honeywell UK").all()

    hw_de = df2[df2["Contact ID"] == "honeywell.de@fleetlogistics.com"]
    assert not hw_de.empty
    assert (hw_de["client"] == "Honeywell International").all()


def test_tally_counts_within_and_violated_case_insensitively():
    df = _load_fixture()
    df = processor.filter_services(df, ["Updates"])
    df = processor.assign_client(df, _clients())
    tallies = processor.tally(df)
    # Row 1006 is a Honeywell UK SLA Violated in services; 1001 and 1009
    # are within SLA BUT 1009 has a blank status and should be ignored.
    # Note row 1006 Subject contains "callback" — in real flow it would be
    # split out; here we test tally itself without the callback filter.
    assert "Honeywell UK" in tallies
    w, o = tallies["Honeywell UK"]
    assert w == 1  # 1001
    assert o == 1  # 1006


def test_tally_ignores_blank_resolution_status():
    df = _load_fixture()
    df = processor.filter_services(df, ["Updates"])
    df = processor.assign_client(df, _clients())
    # Row 1009 (blank status, Honeywell UK) should not bump either count
    tallies = processor.tally(df)
    w, o = tallies["Honeywell UK"]
    assert w + o == 2  # 1001 and 1006 only


def test_callback_filter_picks_subject_match():
    df = _load_fixture()
    df = processor.filter_services(df, ["Updates"])
    rule = {"field": "Subject", "match": ".*call\\s*back.*"}
    callbacks = processor.filter_callbacks(df, rule)
    # 1005 "Callback request" and 1006 "...callback..." match
    assert set(callbacks["Ticket ID"]) == {"1005", "1006"}


def test_summarise_zero_ticket_template_clients_get_zeros():
    df = _load_fixture()
    df_svc = processor.filter_services(df, ["Updates"])
    df_svc = processor.assign_client(df_svc, _clients())
    summary = processor.summarise(
        df_services=df_svc,
        df_callbacks=df_svc.iloc[0:0].assign(client=pd.Series(dtype=object)),
        template_clients=["Honeywell UK", "Ashfield Healthcare", "Beko"],
        month_label="March 2026",
    )
    # Beko had no tickets, must still appear with (0,0)
    assert summary.services["Beko"] == (0, 0)
    assert summary.services["Honeywell UK"][0] >= 1


def test_summarise_unmapped_senders_reported():
    df = _load_fixture()
    df_svc = processor.filter_services(df, ["Updates"])
    df_svc = processor.assign_client(df_svc, _clients())
    summary = processor.summarise(
        df_services=df_svc,
        df_callbacks=df_svc.iloc[0:0].assign(client=pd.Series(dtype=object)),
        template_clients=["Honeywell UK"],
        month_label="March 2026",
    )
    # Two unmapped rows in the fixture, different senders
    assert summary.unmapped_total == 2


def test_detect_month_label_returns_dominant_month():
    df = _load_fixture()
    assert processor.detect_month_label(df) == "March 2026"


def test_detect_month_label_handles_empty():
    assert processor.detect_month_label(pd.DataFrame()) == ""
