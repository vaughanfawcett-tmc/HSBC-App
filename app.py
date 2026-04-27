"""HSBC SLA Report — Streamlit entrypoint.

Three screens, one thing on screen at a time:
    1. Upload         - drop the Fresh CRM export
    2. Review         - month, totals, unmapped warnings, approve/download
    3. Confirmation   - sent to Andrew, message id, start over
"""
from __future__ import annotations

import hmac
import io
import os
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

import audit
import mailer
import processor
import template_writer

ROOT = Path(__file__).parent
SETTINGS_PATH = ROOT / "config" / "settings.yaml"
CLIENTS_PATH = ROOT / "config" / "clients.yaml"

# ---------- page config ----------
st.set_page_config(
    page_title="HSBC SLA Report",
    page_icon=None,
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ---------- global style ----------
CUSTOM_CSS = """
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">

<style>
/* ---------- Streamlit chrome off ---------- */
#MainMenu, header, footer, [data-testid="stToolbar"],
[data-testid="stStatusWidget"], [data-testid="stDecoration"] {
    visibility: hidden !important;
    display: none !important;
}
.stDeployButton, .viewerBadge_container__1QSob {display:none !important;}

/* ---------- Page canvas: soft off-white, subtle radial tint ---------- */
html, body, .stApp {
    background:
      radial-gradient(1200px 600px at 85% -10%, #EFEEFC 0%, transparent 55%),
      radial-gradient(900px 500px at -10% 110%, #EAF4FF 0%, transparent 55%),
      #FAFAFC !important;
    background-attachment: fixed !important;
}
.main > div {background: transparent !important;}

/* ---------- Typography ---------- */
html, body, [class*="css"], .stApp, .stMarkdown, .stButton > button,
.stDownloadButton > button, input, textarea, select {
    font-family: "Inter", -apple-system, "SF Pro Text", "Helvetica Neue",
                 Helvetica, Arial, sans-serif !important;
    -webkit-font-smoothing: antialiased;
    font-feature-settings: "cv11", "ss01", "ss03";
    color: #111114;
}

/* ---------- Layout ---------- */
.block-container {
    max-width: 760px !important;
    padding-top: 4rem !important;
    padding-bottom: 6rem !important;
}

/* Animation on screen change */
.block-container > div {
    animation: fadeUp 420ms cubic-bezier(0.22, 1, 0.36, 1) both;
}
@keyframes fadeUp {
    from { opacity: 0; transform: translateY(8px); }
    to   { opacity: 1; transform: translateY(0); }
}

/* ---------- Typography scale ---------- */
.hero-eyebrow {
    font-size: 0.8125rem;
    font-weight: 500;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #7A7A86;
    margin: 0 0 1rem 0;
}
.hero-title {
    font-size: 3.25rem;
    font-weight: 600;
    letter-spacing: -0.035em;
    line-height: 1.05;
    margin: 0 0 0.75rem 0;
    color: #0B0B0F;
    background: linear-gradient(180deg, #0B0B0F 0%, #34343D 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}
.hero-sub {
    color: #6E6E78;
    font-size: 1.1875rem;
    line-height: 1.5;
    margin: 0 0 3rem 0;
    font-weight: 400;
}
.section-title {
    font-size: 0.8125rem;
    font-weight: 600;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #8A8A95;
    margin: 2.5rem 0 1rem 0;
}

/* ---------- Step indicator ---------- */
.steps {
    display: flex; gap: 0.5rem; align-items: center;
    margin-bottom: 3rem;
}
.step {
    width: 28px; height: 4px; border-radius: 2px;
    background: #E6E6EC; transition: background 200ms ease;
}
.step.active { background: #0B0B0F; }
.step.done   { background: #0B0B0F; opacity: 0.35; }

/* ---------- Upload drop zone ---------- */
[data-testid="stFileUploader"] {margin-top: 0.5rem !important;}
[data-testid="stFileUploader"] section {
    border: 1.5px dashed #D0D0D9 !important;
    border-radius: 20px !important;
    padding: 3.5rem 2rem !important;
    background: rgba(255,255,255,0.6) !important;
    backdrop-filter: blur(12px) !important;
    -webkit-backdrop-filter: blur(12px) !important;
    transition: all 220ms cubic-bezier(0.22, 1, 0.36, 1);
}
[data-testid="stFileUploader"] section:hover {
    background: rgba(255,255,255,0.85) !important;
    border-color: #0B0B0F !important;
    transform: translateY(-2px);
    box-shadow: 0 20px 40px -20px rgba(11, 11, 15, 0.18);
}
[data-testid="stFileUploader"] section > div:first-child {
    font-size: 1rem !important; color: #111114 !important;
}
[data-testid="stFileUploader"] small {color: #8A8A95 !important;}
[data-testid="stFileUploader"] button {
    background: #0B0B0F !important;
    color: #FFF !important;
    border-radius: 10px !important;
    padding: 0.625rem 1.125rem !important;
    font-weight: 500 !important;
    border: none !important;
    box-shadow: 0 1px 2px rgba(11, 11, 15, 0.15) !important;
    transition: transform 120ms ease;
}
[data-testid="stFileUploader"] button:hover {transform: translateY(-1px);}
[data-testid="stFileUploaderDropzoneInstructions"] span {
    color: #111114 !important; font-weight: 500;
}

/* Upload icon above the drop zone */
.upload-icon {
    width: 72px; height: 72px;
    margin: 0 auto 1.5rem auto;
    display: block;
    color: #0B0B0F;
    opacity: 0.9;
}

/* ---------- Buttons ---------- */
.stButton > button, .stDownloadButton > button {
    min-height: 52px !important;
    border-radius: 12px !important;
    padding: 0.875rem 1.5rem !important;
    font-weight: 500 !important;
    font-size: 0.9375rem !important;
    letter-spacing: -0.005em !important;
    width: 100% !important;
    box-shadow: none !important;
    transition: all 180ms cubic-bezier(0.22, 1, 0.36, 1) !important;
}
.stButton > button[kind="primary"] {
    background: #0B0B0F !important;
    color: #FFF !important;
    border: 1px solid #0B0B0F !important;
    box-shadow: 0 8px 24px -12px rgba(11, 11, 15, 0.45) !important;
}
.stButton > button[kind="primary"]:hover {
    background: #000 !important;
    transform: translateY(-1px);
    box-shadow: 0 12px 28px -10px rgba(11, 11, 15, 0.55) !important;
}
.stButton > button[kind="secondary"], .stDownloadButton > button {
    background: rgba(255,255,255,0.75) !important;
    color: #111114 !important;
    border: 1px solid #E3E3EA !important;
    backdrop-filter: blur(8px) !important;
}
.stButton > button[kind="secondary"]:hover, .stDownloadButton > button:hover {
    background: #FFF !important;
    border-color: #0B0B0F !important;
    transform: translateY(-1px);
}

/* ---------- Hero stat (big % Within) ---------- */
.hero-stat {
    background: rgba(255,255,255,0.7);
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
    border: 1px solid rgba(255,255,255,0.8);
    border-radius: 24px;
    padding: 2.25rem 2.5rem;
    margin: 0 0 1rem 0;
    display: flex;
    align-items: flex-end;
    justify-content: space-between;
    box-shadow: 0 20px 40px -28px rgba(11, 11, 15, 0.18);
}
.hero-stat .pct {
    font-size: 4.75rem;
    font-weight: 600;
    letter-spacing: -0.04em;
    line-height: 1;
    color: #0B0B0F;
    background: linear-gradient(180deg, #0B0B0F 0%, #3A3A45 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}
.hero-stat .pct-unit {
    font-size: 1.625rem; font-weight: 500;
    color: #7A7A86; margin-left: 0.125rem;
    letter-spacing: -0.01em;
}
.hero-stat .pct-label {
    font-size: 0.8125rem; letter-spacing: 0.1em;
    text-transform: uppercase; color: #7A7A86;
    margin-top: 0.5rem; font-weight: 500;
}
.hero-stat .sparkbar {
    flex: 1; margin-left: 2rem;
    height: 64px;
    display: flex; flex-direction: column; justify-content: flex-end;
}
.hero-stat .sparkbar .track {
    height: 8px; background: #EDEDF2; border-radius: 999px; overflow: hidden;
    position: relative;
}
.hero-stat .sparkbar .fill {
    height: 100%;
    background: linear-gradient(90deg, #0B0B0F 0%, #2E2E38 100%);
    border-radius: 999px;
    transition: width 600ms cubic-bezier(0.22, 1, 0.36, 1);
}
.hero-stat .sparkbar-label {
    display: flex; justify-content: space-between;
    font-size: 0.75rem; color: #8A8A95; margin-bottom: 0.5rem;
    font-feature-settings: "tnum";
}

/* ---------- Stat tiles row ---------- */
.tile-row {display:grid; grid-template-columns: repeat(3, 1fr); gap: 0.75rem;}
.tile {
    background: rgba(255,255,255,0.7);
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
    border: 1px solid rgba(255,255,255,0.8);
    border-radius: 16px;
    padding: 1.25rem 1.375rem;
    box-shadow: 0 10px 24px -20px rgba(11, 11, 15, 0.15);
}
.tile .tile-label {
    font-size: 0.75rem; letter-spacing: 0.08em; text-transform: uppercase;
    color: #8A8A95; font-weight: 500; margin: 0;
}
.tile .tile-value {
    font-size: 1.75rem; font-weight: 600; letter-spacing: -0.025em;
    color: #0B0B0F; margin: 0.25rem 0 0 0;
    font-feature-settings: "tnum";
}
.tile.accent-ok .tile-value {color: #0B0B0F;}
.tile.accent-warn .tile-value {color: #B84500;}
.tile .tile-dot {
    width: 6px; height: 6px; border-radius: 999px;
    display: inline-block; margin-right: 0.4375rem;
    vertical-align: middle;
}
.tile.accent-ok   .tile-dot {background: #34C759;}
.tile.accent-warn .tile-dot {background: #FF9F0A;}

/* ---------- Warning card ---------- */
.warn-card {
    background: linear-gradient(180deg, #FFF6E6 0%, #FFF9EE 100%);
    border: 1px solid #FAD38A;
    border-radius: 16px;
    padding: 1.25rem 1.5rem;
    margin: 1.75rem 0;
    display: flex; gap: 1rem;
    box-shadow: 0 10px 24px -20px rgba(184, 69, 0, 0.2);
}
.warn-card .warn-icon {
    flex-shrink: 0; color: #B84500; margin-top: 2px;
}
.warn-card .warn-body {flex: 1;}
.warn-card .warn-title {
    font-weight: 600; color: #7A3700; margin: 0 0 0.375rem 0;
    font-size: 0.9375rem; letter-spacing: -0.005em;
}
.warn-card .warn-text {
    color: #8B4B00; font-size: 0.875rem; line-height: 1.55; margin: 0;
}
.warn-card ul {margin: 0.75rem 0 0 0; padding: 0; list-style: none;}
.warn-card li {
    display: flex; justify-content: space-between;
    padding: 0.375rem 0; border-bottom: 1px solid rgba(184,69,0,0.08);
    font-size: 0.875rem; color: #5D3300;
}
.warn-card li:last-child {border: none;}
.warn-card li code {
    font-family: "JetBrains Mono", ui-monospace, monospace;
    background: rgba(255,255,255,0.6); padding: 2px 6px; border-radius: 6px;
    font-size: 0.8125rem; color: #5D3300;
}
.warn-card li .count {
    font-variant-numeric: tabular-nums; font-weight: 500; color: #7A3700;
}

/* ---------- Error card ---------- */
.error-card {
    background: #FFF0ED; border: 1px solid #F5A599;
    color: #6E1500; border-radius: 14px;
    padding: 1rem 1.25rem; margin: 1.25rem 0;
    font-size: 0.9375rem; line-height: 1.55;
    display: flex; gap: 0.75rem; align-items: flex-start;
}
.error-card .err-icon {flex-shrink: 0; color: #CC2B00; margin-top: 1px;}

/* ---------- Dataframe styling ---------- */
[data-testid="stDataFrame"] {
    border: 1px solid #E8E8EE !important;
    border-radius: 14px !important;
    overflow: hidden !important;
    background: rgba(255,255,255,0.7) !important;
    backdrop-filter: blur(12px) !important;
    box-shadow: 0 10px 24px -20px rgba(11, 11, 15, 0.12) !important;
}
[data-testid="stDataFrame"] thead tr th {
    background: rgba(246,246,250,0.9) !important;
    border-bottom: 1px solid #E8E8EE !important;
    font-weight: 600 !important;
    color: #4A4A54 !important;
    font-size: 0.8125rem !important;
    letter-spacing: -0.005em !important;
}
[data-testid="stDataFrame"] tbody tr td {
    font-size: 0.9375rem !important;
    color: #1D1D26 !important;
    border-color: #F2F2F5 !important;
}

/* ---------- Device code block ---------- */
.devcode {
    background: linear-gradient(145deg, #0B0B0F 0%, #1E1E28 100%);
    color: #F5F5FA;
    font-family: "JetBrains Mono", ui-monospace, monospace;
    padding: 1.5rem 1.75rem; border-radius: 16px;
    margin: 1.25rem 0; line-height: 1.7;
    font-size: 0.9375rem;
    box-shadow: 0 20px 40px -20px rgba(11, 11, 15, 0.4);
}

/* ---------- Test-mode banner ---------- */
.test-banner {
    background: linear-gradient(180deg, #EAF4FF 0%, #F1F8FF 100%);
    border: 1px solid #B8D9F5;
    color: #00335C;
    padding: 0.75rem 1.125rem; border-radius: 12px;
    margin-bottom: 2rem;
    font-size: 0.875rem;
    display: flex; align-items: center; gap: 0.625rem;
    letter-spacing: -0.005em;
}
.test-banner svg {flex-shrink: 0; color: #0066CC;}

/* ---------- Confirmation screen ---------- */
.confirm-wrap {
    text-align: center;
    margin-top: 5rem;
    animation: popIn 520ms cubic-bezier(0.22, 1, 0.36, 1) both;
}
@keyframes popIn {
    from {opacity:0; transform: scale(0.92);}
    to   {opacity:1; transform: scale(1);}
}
.confirm-icon {
    width: 96px; height: 96px;
    margin: 0 auto 2rem auto;
    display: block;
    color: #0B0B0F;
    filter: drop-shadow(0 12px 20px rgba(11,11,15,0.18));
}
.confirm-title {
    font-size: 2.75rem;
    font-weight: 600;
    letter-spacing: -0.035em;
    color: #0B0B0F;
    margin: 0 0 0.5rem 0;
}
.confirm-recipient {
    font-family: "JetBrains Mono", ui-monospace, monospace;
    color: #4A4A54; font-size: 1rem;
}
.confirm-meta {
    color: #8A8A95; font-size: 0.8125rem;
    margin-top: 1.25rem; line-height: 1.7;
    font-family: "JetBrains Mono", ui-monospace, monospace;
}

/* Streamlit tweaks */
div[data-testid="stHorizontalBlock"] {gap: 0.75rem !important;}
div[data-baseweb="notification"] {border-radius: 12px !important;}

/* Success green for low outside tile */
.tile.accent-ok .tile-label {color: #6E6E78;}

/* Subtle divider */
.soft-divider {
    height: 1px;
    background: linear-gradient(90deg, transparent, #E8E8EE 20%, #E8E8EE 80%, transparent);
    margin: 2.5rem 0 1.5rem 0;
}
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ---------- SVG icon library ----------
def svg_upload_cloud() -> str:
    return (
        '<svg class="upload-icon" xmlns="http://www.w3.org/2000/svg" '
        'viewBox="0 0 24 24" fill="none" stroke="currentColor" '
        'stroke-width="1.25" stroke-linecap="round" stroke-linejoin="round">'
        '<path d="M12 13v9"/>'
        '<path d="m16 17-4-4-4 4"/>'
        '<path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"/>'
        '</svg>'
    )


def svg_warning() -> str:
    return (
        '<svg class="warn-icon" xmlns="http://www.w3.org/2000/svg" '
        'width="22" height="22" viewBox="0 0 24 24" fill="none" '
        'stroke="currentColor" stroke-width="1.75" stroke-linecap="round" '
        'stroke-linejoin="round">'
        '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/>'
        '<path d="M12 9v4"/><path d="M12 17h.01"/>'
        '</svg>'
    )


def svg_error() -> str:
    return (
        '<svg class="err-icon" xmlns="http://www.w3.org/2000/svg" '
        'width="20" height="20" viewBox="0 0 24 24" fill="none" '
        'stroke="currentColor" stroke-width="1.75" stroke-linecap="round" '
        'stroke-linejoin="round">'
        '<circle cx="12" cy="12" r="10"/><path d="m15 9-6 6"/><path d="m9 9 6 6"/>'
        '</svg>'
    )


def svg_info() -> str:
    return (
        '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" '
        'viewBox="0 0 24 24" fill="none" stroke="currentColor" '
        'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
        '<circle cx="12" cy="12" r="10"/><path d="M12 16v-4"/><circle cx="12" cy="8" r=".5" fill="currentColor"/>'
        '</svg>'
    )


def svg_check() -> str:
    return (
        '<svg class="confirm-icon" xmlns="http://www.w3.org/2000/svg" '
        'viewBox="0 0 24 24" fill="none" stroke="currentColor" '
        'stroke-width="1.25" stroke-linecap="round" stroke-linejoin="round">'
        '<circle cx="12" cy="12" r="10"/>'
        '<path d="m9 12 2 2 4-4"/>'
        '</svg>'
    )


def render_steps(active_index: int) -> str:
    classes = []
    for i in range(3):
        if i < active_index:
            classes.append("step done")
        elif i == active_index:
            classes.append("step active")
        else:
            classes.append("step")
    return (
        "<div class='steps'>"
        + "".join(f"<span class='{c}'></span>" for c in classes)
        + "</div>"
    )


# ---------- session state ----------
def _init_state() -> None:
    defaults = {
        "screen": "upload",
        "source_bytes": None,
        "source_name": None,
        "source_sha": None,
        "summary": None,
        "output_path": None,
        "output_bytes": None,
        "output_sha": None,
        "month_label": "",
        "upload_error": None,
        "send_error": None,
        "send_result": None,
        "device_flow_msg": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


_init_state()


# ---------- helpers ----------
REQUIRED_COLS = {"Resolution status", "Group", "Created time", "Full name", "Contact ID", "Subject"}


def _load_settings():
    return processor.load_settings(str(SETTINGS_PATH))


def _process(csv_bytes: bytes, settings: dict):
    df = pd.read_csv(io.BytesIO(csv_bytes), dtype=str, keep_default_na=False)
    missing = REQUIRED_COLS - set(df.columns)
    if missing:
        raise ValueError(
            "That doesn't look like a Fresh export. Missing columns: "
            + ", ".join(sorted(missing))
        )

    clients = processor.load_clients_config(str(CLIENTS_PATH))
    template_path = ROOT / settings["template_path"]
    template_clients = template_writer.read_template_clients(template_path)

    exclude_groups = settings.get("filters", {}).get("exclude_groups", []) or []
    callback_rule = settings.get("filters", {}).get("callback_rule") or None
    identifier_field = settings.get("identifier_field", "Full name")

    df_services = processor.filter_services(df, exclude_groups)
    df_callbacks = processor.filter_callbacks(df_services, callback_rule)

    # Callback rows should NOT also be counted in services totals.
    if not df_callbacks.empty:
        df_services = df_services.drop(index=df_callbacks.index)

    df_services = processor.assign_client(df_services, clients)
    df_callbacks = processor.assign_client(df_callbacks, clients)

    month_label = processor.detect_month_label(df) or datetime.utcnow().strftime("%B %Y")

    summary = processor.summarise(
        df_services=df_services,
        df_callbacks=df_callbacks,
        template_clients=template_clients,
        month_label=month_label,
        identifier_field=identifier_field,
    )
    return summary, template_path


def _build_xlsx(summary, template_path: Path, settings: dict) -> tuple[Path, bytes]:
    out_dir = ROOT / settings.get("output_dir", "output")
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    safe_month = summary.month_label.replace(" ", "_") or "report"
    out_path = out_dir / f"HSBC_SLA_{safe_month}_{stamp}.xlsx"
    template_writer.fill_template(
        template_path=template_path,
        output_path=out_path,
        services_tallies=summary.services,
        callbacks_tallies=summary.callbacks,
        month_label=summary.month_label,
    )
    return out_path, out_path.read_bytes()


def _go(screen: str) -> None:
    st.session_state.screen = screen


def _reset() -> None:
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    _init_state()


# ---------- screen 1: upload ----------
def screen_upload() -> None:
    st.markdown(render_steps(0), unsafe_allow_html=True)
    st.markdown(
        "<p class='hero-eyebrow'>HSBC · Monthly report</p>"
        "<h1 class='hero-title'>SLA report</h1>"
        "<p class='hero-sub'>Drop your Fresh CRM export below. "
        "We'll tally, fill the template, and hand it back for review.</p>",
        unsafe_allow_html=True,
    )
    st.markdown(svg_upload_cloud(), unsafe_allow_html=True)

    uploaded = st.file_uploader(
        " ",
        type=["csv"],
        accept_multiple_files=False,
        label_visibility="collapsed",
        key="uploader",
    )

    if st.session_state.upload_error:
        st.markdown(
            f"<div class='error-card'>{svg_error()}"
            f"<div>{st.session_state.upload_error}</div></div>",
            unsafe_allow_html=True,
        )

    if uploaded is not None:
        try:
            settings = _load_settings()
            raw = uploaded.getvalue()
            summary, template_path = _process(raw, settings)
            out_path, out_bytes = _build_xlsx(summary, template_path, settings)
            st.session_state.source_bytes = raw
            st.session_state.source_name = uploaded.name
            st.session_state.source_sha = audit.sha256_bytes(raw)
            st.session_state.summary = summary
            st.session_state.month_label = summary.month_label
            st.session_state.output_path = str(out_path)
            st.session_state.output_bytes = out_bytes
            st.session_state.output_sha = audit.sha256_bytes(out_bytes)
            st.session_state.upload_error = None
            _go("review")
            st.rerun()
        except Exception as e:  # noqa: BLE001
            st.session_state.upload_error = str(e)
            st.rerun()


# ---------- screen 2: review ----------
def screen_review() -> None:
    summary = st.session_state.summary
    settings = _load_settings()
    email_cfg = settings.get("email", {})
    test_mode = bool(email_cfg.get("test_mode"))
    test_recipient = email_cfg.get("test_recipient") or ""

    st.markdown(render_steps(1), unsafe_allow_html=True)

    if test_mode:
        tgt = test_recipient or "(test_recipient not set)"
        st.markdown(
            f"<div class='test-banner'>{svg_info()}"
            f"<span><strong>Test mode.</strong> Mail will go to "
            f"<code>{tgt}</code>, not the HSBC contact.</span></div>",
            unsafe_allow_html=True,
        )

    svc_w, svc_o = summary.services_totals
    cb_w, cb_o = summary.callbacks_totals
    total_within = svc_w + cb_w
    total_outside = svc_o + cb_o
    total_tickets = total_within + total_outside
    pct = (total_within / total_tickets * 100) if total_tickets else 0.0

    st.markdown(
        f"<p class='hero-eyebrow'>{summary.month_label or 'Report'}</p>"
        f"<h1 class='hero-title'>Looks good.</h1>"
        f"<p class='hero-sub'>{total_tickets:,} ticket{'s' if total_tickets != 1 else ''} "
        f"processed across {len([c for c,v in summary.services.items() if v[0]+v[1]])} "
        f"active client{'s' if len([c for c,v in summary.services.items() if v[0]+v[1]]) != 1 else ''}. "
        f"Review and send when you're happy.</p>",
        unsafe_allow_html=True,
    )

    # Hero % Within + bar
    st.markdown(
        f"<div class='hero-stat'>"
        f"<div>"
        f"<div><span class='pct'>{pct:.1f}</span><span class='pct-unit'>%</span></div>"
        f"<p class='pct-label'>Within SLA</p>"
        f"</div>"
        f"<div class='sparkbar'>"
        f"<div class='sparkbar-label'>"
        f"<span>{total_within:,} within</span>"
        f"<span>{total_outside:,} outside</span>"
        f"</div>"
        f"<div class='track'><div class='fill' style='width:{pct:.2f}%'></div></div>"
        f"</div>"
        f"</div>",
        unsafe_allow_html=True,
    )

    # Three tiles: services / callbacks / unmapped
    svc_total = svc_w + svc_o
    cb_total = cb_w + cb_o
    unmapped_accent = "accent-warn" if summary.unmapped_total else "accent-ok"
    unmapped_dot = "accent-warn" if summary.unmapped_total else "accent-ok"
    st.markdown(
        f"<div class='tile-row'>"
        f"<div class='tile accent-ok'>"
        f"<p class='tile-label'><span class='tile-dot'></span>Services</p>"
        f"<p class='tile-value'>{svc_total:,}</p></div>"
        f"<div class='tile accent-ok'>"
        f"<p class='tile-label'><span class='tile-dot'></span>Call backs</p>"
        f"<p class='tile-value'>{cb_total:,}</p></div>"
        f"<div class='tile {unmapped_accent}'>"
        f"<p class='tile-label'><span class='tile-dot'></span>Unmapped</p>"
        f"<p class='tile-value'>{summary.unmapped_total:,}</p></div>"
        f"</div>",
        unsafe_allow_html=True,
    )

    # Unmapped warning
    if summary.unmapped_senders:
        top_items = sorted(summary.unmapped_senders.items(), key=lambda x: -x[1])[:15]
        rows = "".join(
            f"<li><code>{name}</code><span class='count'>{count:,}</span></li>"
            for name, count in top_items
        )
        more = ""
        extra = len(summary.unmapped_senders) - 15
        if extra > 0:
            more = f"<li style='color:#8B4B00;justify-content:center;'>+ {extra:,} more sender{'s' if extra != 1 else ''}</li>"
        st.markdown(
            f"<div class='warn-card'>"
            f"{svg_warning()}"
            f"<div class='warn-body'>"
            f"<p class='warn-title'>{len(summary.unmapped_senders):,} unmapped sender"
            f"{'s' if len(summary.unmapped_senders) != 1 else ''} "
            f"· {summary.unmapped_total:,} ticket"
            f"{'s' if summary.unmapped_total != 1 else ''}</p>"
            f"<p class='warn-text'>These senders didn't match any client in "
            f"<code>config/clients.yaml</code>. Add or widen a rule to include "
            f"them, or continue if they should be ignored.</p>"
            f"<ul>{rows}{more}</ul>"
            f"</div></div>",
            unsafe_allow_html=True,
        )

    for w in summary.warnings:
        st.markdown(
            f"<div class='warn-card'>{svg_warning()}"
            f"<div class='warn-body'><p class='warn-text'>{w}</p></div></div>",
            unsafe_allow_html=True,
        )

    # Per-client table with inline progress column
    st.markdown("<p class='section-title'>Per client</p>", unsafe_allow_html=True)
    rows = []
    for client in summary.services:
        sw, so = summary.services[client]
        cw, co = summary.callbacks.get(client, (0, 0))
        total = sw + so + cw + co
        # ProgressColumn uses the raw value for both the bar fill and the
        # printf-formatted label, so we store it as 0-100 (not 0-1).
        pct_within = ((sw + cw) / total * 100) if total else None
        rows.append(
            {
                "Client": client,
                "Total": total,
                "Within": sw + cw,
                "Outside": so + co,
                "% Within": pct_within,
            }
        )
    df_display = pd.DataFrame(rows).sort_values(by="Total", ascending=False)

    st.dataframe(
        df_display,
        use_container_width=True,
        hide_index=True,
        height=min(480, 56 + 35 * len(df_display)),
        column_config={
            "Client": st.column_config.TextColumn("Client", width="medium"),
            "Total": st.column_config.NumberColumn("Total", format="%d", width="small"),
            "Within": st.column_config.NumberColumn("Within", format="%d", width="small"),
            "Outside": st.column_config.NumberColumn("Outside", format="%d", width="small"),
            "% Within": st.column_config.ProgressColumn(
                "% Within",
                format="%.0f%%",
                min_value=0.0,
                max_value=100.0,
                width="medium",
            ),
        },
    )

    # Send error
    if st.session_state.send_error:
        st.markdown(
            f"<div class='error-card'>{svg_error()}"
            f"<div>{st.session_state.send_error}</div></div>",
            unsafe_allow_html=True,
        )

    # Device code
    if st.session_state.device_flow_msg:
        st.markdown(
            f"<div class='devcode'>{st.session_state.device_flow_msg}</div>",
            unsafe_allow_html=True,
        )

    st.markdown("<div class='soft-divider'></div>", unsafe_allow_html=True)

    # Actions
    left, right = st.columns(2)
    with left:
        st.download_button(
            "Download for review",
            data=st.session_state.output_bytes,
            file_name=Path(st.session_state.output_path).name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_xlsx",
            type="secondary",
            use_container_width=True,
        )
    with right:
        if st.button("Approve & send", key="send_btn", type="primary", use_container_width=True):
            _do_send(settings)


def _do_send(settings: dict) -> None:
    st.session_state.send_error = None
    email_cfg = settings.get("email", {})
    recipient = email_cfg.get("recipient") or ""
    if email_cfg.get("test_mode"):
        tr = email_cfg.get("test_recipient") or ""
        if not tr:
            st.session_state.send_error = (
                "Test mode is on but test_recipient is empty in settings.yaml."
            )
            st.rerun()
            return
        recipient = tr
    cc = email_cfg.get("cc") or []

    subject = (email_cfg.get("subject_template") or "").format(
        month_label=st.session_state.month_label
    )
    body = (email_cfg.get("body_template") or "").format(
        month_label=st.session_state.month_label
    )

    def _show_code(msg: str, _flow: dict) -> None:
        st.session_state.device_flow_msg = msg

    try:
        token = mailer.acquire_token(interactive_callback=_show_code)
        expected_sender = (email_cfg.get("expected_sender_email") or "").strip().lower()
        if expected_sender:
            me = mailer.current_user(token)
            who = (me.get("mail") or me.get("userPrincipalName") or "").strip().lower()
            if who != expected_sender:
                st.session_state.send_error = (
                    f"Signed in as {who or '(unknown)'}, but this app is "
                    f"configured to send only as {expected_sender}. Sign out, "
                    f"then sign back in as {expected_sender}."
                )
                st.session_state.device_flow_msg = None
                st.rerun()
                return
        result = mailer.send_mail(
            token=token,
            recipient=recipient,
            subject=subject,
            body_text=body,
            attachment_path=st.session_state.output_path,
            cc=cc,
        )
    except mailer.MailerError as e:
        st.session_state.send_error = str(e)
        st.rerun()
        return
    except Exception as e:  # noqa: BLE001
        st.session_state.send_error = f"Unexpected error: {e}"
        st.rerun()
        return

    # Audit log
    summary = st.session_state.summary
    svc_w, svc_o = summary.services_totals
    cb_w, cb_o = summary.callbacks_totals
    audit_path_override = os.environ.get("HSBC_SLA_AUDIT_LOG_PATH")
    audit_path = Path(audit_path_override) if audit_path_override else (
        ROOT / settings.get("audit_log_path", "audit_log.csv")
    )
    audit.append_entry(
        audit_path,
        {
            "user_email": result.user_email,
            "source_csv_sha256": st.session_state.source_sha or "",
            "output_xlsx_sha256": st.session_state.output_sha or "",
            "month_label": summary.month_label,
            "services_total_within": svc_w,
            "services_total_outside": svc_o,
            "callbacks_total_within": cb_w,
            "callbacks_total_outside": cb_o,
            "unmapped_count": summary.unmapped_total,
            "recipient": recipient,
            "graph_message_id": result.message_id,
        },
    )
    st.session_state.send_result = {
        "message_id": result.message_id,
        "recipient": recipient,
        "when": datetime.now().strftime("%d %b %Y, %H:%M"),
    }
    st.session_state.device_flow_msg = None
    _go("done")
    st.rerun()


# ---------- screen 3: confirmation ----------
def screen_done() -> None:
    r = st.session_state.send_result or {}
    recipient = r.get("recipient", "recipient")
    when = r.get("when", "")
    msg_id = r.get("message_id", "")

    st.markdown(
        f"<div class='confirm-wrap'>"
        f"{svg_check()}"
        f"<p class='confirm-title'>Sent.</p>"
        f"<p class='confirm-recipient'>{recipient}</p>"
        f"<p class='confirm-meta'>{when}<br>Ref · {msg_id}</p>"
        f"</div>",
        unsafe_allow_html=True,
    )
    # Center the button
    _, col, _ = st.columns([1, 2, 1])
    with col:
        if st.button("Start new report", type="primary", use_container_width=True):
            _reset()
            st.rerun()


# ---------- access gate ----------
def _check_password() -> bool:
    """Block the app behind a shared password set via APP_PASSWORD env var.
    Returns True when the visitor is authenticated. When APP_PASSWORD is
    unset (typical for local dev) the gate is open."""
    expected = os.environ.get("APP_PASSWORD", "")
    if not expected:
        return True
    if st.session_state.get("auth_ok"):
        return True

    st.markdown(
        "<p class='hero-eyebrow'>HSBC · Internal</p>"
        "<h1 class='hero-title'>SLA report</h1>"
        "<p class='hero-sub'>Sign in with the shared password to continue.</p>",
        unsafe_allow_html=True,
    )
    pwd = st.text_input(
        "Password", type="password", key="pwd_input", label_visibility="collapsed",
        placeholder="Password",
    )
    if st.button("Continue", type="primary", use_container_width=True, key="auth_btn"):
        if hmac.compare_digest(pwd, expected):
            st.session_state.auth_ok = True
            st.session_state.pop("auth_error", None)
            st.rerun()
        else:
            st.session_state.auth_error = "Incorrect password."
            st.rerun()
    if st.session_state.get("auth_error"):
        st.markdown(
            f"<div class='error-card'>{svg_error()}"
            f"<div>{st.session_state.auth_error}</div></div>",
            unsafe_allow_html=True,
        )
    return False


if not _check_password():
    st.stop()


# ---------- router ----------
screen = st.session_state.screen
if screen == "review":
    screen_review()
elif screen == "done":
    screen_done()
else:
    screen_upload()
