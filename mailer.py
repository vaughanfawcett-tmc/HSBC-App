"""Microsoft Graph email sender.

Uses MSAL device-code flow so Bethany signs in once in a browser; the
resulting refresh token is cached to ~/.hsbc_sla_tool/token.bin and reused
on subsequent runs. Mail is sent as her via /me/sendMail, so the recipient
sees it from her mailbox, not from a service account or bot.

Swap this module for a Gmail equivalent if Miles is on Google Workspace;
the fill_and_send contract the UI depends on stays the same.
"""
from __future__ import annotations

import base64
import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import msal
import requests

AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["Mail.Send", "User.Read"]
GRAPH_SENDMAIL = "https://graph.microsoft.com/v1.0/me/sendMail"
GRAPH_ME = "https://graph.microsoft.com/v1.0/me"

# Microsoft's public "Graph Explorer"-style client ID works for personal-ish
# device-code flows without needing an app registration. For a regulated
# production deployment Miles should register their own app and override
# this via the HSBC_SLA_CLIENT_ID env var.
DEFAULT_CLIENT_ID = os.environ.get(
    "HSBC_SLA_CLIENT_ID", "14d82eec-204b-4c2f-b7e8-296a70dab67e"
)

TOKEN_DIR = Path(os.environ.get("HSBC_SLA_TOKEN_DIR", Path.home() / ".hsbc_sla_tool"))
TOKEN_FILE = TOKEN_DIR / "token.bin"


class MailerError(RuntimeError):
    pass


@dataclass
class SendResult:
    message_id: str
    user_email: str


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if TOKEN_FILE.exists():
        try:
            cache.deserialize(TOKEN_FILE.read_text())
        except Exception:
            # Corrupt cache — start fresh rather than block the user.
            pass
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        TOKEN_DIR.mkdir(parents=True, exist_ok=True)
        TOKEN_FILE.write_text(cache.serialize())
        try:
            os.chmod(TOKEN_FILE, 0o600)
        except OSError:
            pass


def _build_app(cache: msal.SerializableTokenCache, client_id: str = DEFAULT_CLIENT_ID) -> msal.PublicClientApplication:
    return msal.PublicClientApplication(
        client_id=client_id,
        authority=AUTHORITY,
        token_cache=cache,
    )


def acquire_token(interactive_callback=None, client_id: str = DEFAULT_CLIENT_ID) -> str:
    """Return a valid access token. Uses cached refresh token silently when
    possible; otherwise triggers device-code flow and reports the code via
    interactive_callback(msg) so the Streamlit UI can display it."""
    cache = _load_cache()
    app = _build_app(cache, client_id=client_id)

    accounts = app.get_accounts()
    result: dict[str, Any] | None = None
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise MailerError(f"Failed to start device-code flow: {flow}")
        if interactive_callback is not None:
            interactive_callback(flow.get("message", ""), flow)
        else:
            print(flow.get("message", ""))
        result = app.acquire_token_by_device_flow(flow)

    _save_cache(cache)

    if "access_token" not in result:
        raise MailerError(
            f"Sign-in failed: {result.get('error_description') or result}"
        )
    return result["access_token"]


def current_user(token: str) -> dict[str, Any]:
    resp = requests.get(
        GRAPH_ME,
        headers={"Authorization": f"Bearer {token}"},
        timeout=30,
    )
    if resp.status_code >= 400:
        raise MailerError(f"Graph /me failed: {resp.status_code} {resp.text}")
    return resp.json()


def send_mail(
    token: str,
    recipient: str,
    subject: str,
    body_text: str,
    attachment_path: str | Path,
    cc: list[str] | None = None,
) -> SendResult:
    """Send a plain-text email with the xlsx attached. Returns the message id
    reported by Graph (from the Location header, when available)."""
    cc = cc or []
    attachment_path = Path(attachment_path)
    if not attachment_path.exists():
        raise MailerError(f"Attachment not found: {attachment_path}")
    payload_bytes = attachment_path.read_bytes()
    attachment_b64 = base64.b64encode(payload_bytes).decode("ascii")

    me = current_user(token)
    user_email = me.get("mail") or me.get("userPrincipalName") or ""

    message = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body_text},
            "toRecipients": [{"emailAddress": {"address": recipient}}],
            "ccRecipients": [{"emailAddress": {"address": c}} for c in cc],
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": attachment_path.name,
                    "contentType": (
                        "application/vnd.openxmlformats-officedocument"
                        ".spreadsheetml.sheet"
                    ),
                    "contentBytes": attachment_b64,
                }
            ],
        },
        "saveToSentItems": True,
    }

    resp = requests.post(
        GRAPH_SENDMAIL,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        data=json.dumps(message),
        timeout=60,
    )

    if resp.status_code not in (200, 202):
        raise MailerError(
            f"sendMail failed ({resp.status_code}): {resp.text}"
        )

    # Graph's sendMail returns 202 Accepted with no body and no message id.
    # Pull the client-request-id header for traceability.
    msg_id = resp.headers.get("client-request-id") or resp.headers.get("request-id") or "(accepted)"
    return SendResult(message_id=msg_id, user_email=user_email)


def sign_out() -> None:
    """Remove cached credentials."""
    if TOKEN_FILE.exists():
        TOKEN_FILE.unlink()
