"""
email_handler.py — Microsoft Graph API integration for VRCOMM LINE Bot
-----------------------------------------------------------------------
Handles:
  - OAuth2 client credentials token
  - Webhook subscription create / renew
  - Fetch full email content
  - Send email reply (threaded)

Required env vars:
  MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, MS_EMAIL_ADDRESS
"""

import os, json, logging, requests
from datetime import datetime, timezone, timedelta

logger = logging.getLogger(__name__)

TENANT_ID     = os.environ.get("MS_TENANT_ID", "")
CLIENT_ID     = os.environ.get("MS_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("MS_CLIENT_SECRET", "")
EMAIL_ADDRESS = os.environ.get("MS_EMAIL_ADDRESS", "")
WEBHOOK_URL   = os.environ.get("WEBHOOK_BASE_URL", "") + "/email-webhook"
CLIENT_STATE  = "vrcomm-graph-secret-2026"   # verified on each notification

GRAPH_BASE    = "https://graph.microsoft.com/v1.0"
TOKEN_URL     = "https://login.microsoftonline.com/%s/oauth2/v2.0/token" % TENANT_ID

_access_token = None
_token_expiry = None


# ── Authentication ────────────────────────────────────────────────────────────

def get_access_token() -> str:
    """Return a valid access token, refreshing if expired."""
    global _access_token, _token_expiry
    now = datetime.now(timezone.utc)
    if _access_token and _token_expiry and now < _token_expiry:
        return _access_token
    try:
        resp = requests.post(TOKEN_URL, data={
            "grant_type":    "client_credentials",
            "client_id":     CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "scope":         "https://graph.microsoft.com/.default",
        }, timeout=15)
        resp.raise_for_status()
        data = resp.json()
        _access_token = data["access_token"]
        _token_expiry = now + timedelta(seconds=data.get("expires_in", 3600) - 60)
        logger.info("Graph token refreshed, expires: %s", _token_expiry)
        return _access_token
    except Exception as e:
        logger.error("Graph token error: %s", e)
        raise


def _headers() -> dict:
    return {
        "Authorization": "Bearer %s" % get_access_token(),
        "Content-Type":  "application/json",
    }


# ── Webhook subscription ──────────────────────────────────────────────────────

def create_subscription() -> dict:
    """
    Create a Graph webhook subscription for the inbox.
    Returns the subscription dict (contains 'id' and 'expirationDateTime').
    """
    expiry = (datetime.now(timezone.utc) + timedelta(hours=71)).strftime(
        "%Y-%m-%dT%H:%M:%S.0000000Z"
    )
    payload = {
        "changeType":          "created",
        "notificationUrl":     WEBHOOK_URL,
        "resource":            "users/%s/mailFolders/inbox/messages" % EMAIL_ADDRESS,
        "expirationDateTime":  expiry,
        "clientState":         CLIENT_STATE,
    }
    resp = requests.post(
        "%s/subscriptions" % GRAPH_BASE,
        headers=_headers(),
        json=payload,
        timeout=15,
    )
    if resp.status_code == 201:
        sub = resp.json()
        logger.info("Graph subscription created: %s exp: %s",
                    sub["id"], sub["expirationDateTime"])
        return sub
    else:
        logger.error("Subscription create failed %d: %s", resp.status_code, resp.text)
        resp.raise_for_status()


def renew_subscription(subscription_id: str) -> dict:
    """Extend an existing subscription by 71 hours."""
    expiry = (datetime.now(timezone.utc) + timedelta(hours=71)).strftime(
        "%Y-%m-%dT%H:%M:%S.0000000Z"
    )
    resp = requests.patch(
        "%s/subscriptions/%s" % (GRAPH_BASE, subscription_id),
        headers=_headers(),
        json={"expirationDateTime": expiry},
        timeout=15,
    )
    if resp.status_code == 200:
        sub = resp.json()
        logger.info("Graph subscription renewed: %s exp: %s",
                    sub["id"], sub["expirationDateTime"])
        return sub
    else:
        logger.error("Subscription renew failed %d: %s", resp.status_code, resp.text)
        resp.raise_for_status()


def list_subscriptions() -> list:
    """Return all active subscriptions for this app."""
    resp = requests.get(
        "%s/subscriptions" % GRAPH_BASE,
        headers=_headers(),
        timeout=15,
    )
    resp.raise_for_status()
    return resp.json().get("value", [])


# ── Email fetch ───────────────────────────────────────────────────────────────

def get_email(message_id: str) -> dict:
    """
    Fetch full email content by message ID.
    Returns dict with keys: id, subject, sender_name, sender_email,
    body_text, body_preview, received_at, internet_message_id
    """
    url = "%s/users/%s/messages/%s" % (GRAPH_BASE, EMAIL_ADDRESS, message_id)
    params = {
        "$select": ("id,subject,sender,from,receivedDateTime,"
                    "bodyPreview,body,internetMessageId,conversationId")
    }
    resp = requests.get(url, headers=_headers(), params=params, timeout=15)
    resp.raise_for_status()
    msg = resp.json()

    sender = msg.get("from", msg.get("sender", {})).get("emailAddress", {})
    body   = msg.get("body", {})

    # Strip HTML tags from body if content type is HTML
    body_text = body.get("content", "")
    if body.get("contentType", "").lower() == "html":
        import re
        body_text = re.sub(r"<[^>]+>", " ", body_text)
        body_text = re.sub(r"\s+", " ", body_text).strip()

    return {
        "id":                  msg.get("id", ""),
        "subject":             msg.get("subject", "(no subject)"),
        "sender_name":         sender.get("name", ""),
        "sender_email":        sender.get("address", ""),
        "body_text":           body_text[:4000],   # cap at 4000 chars for Claude
        "body_preview":        msg.get("bodyPreview", "")[:500],
        "received_at":         msg.get("receivedDateTime", ""),
        "internet_message_id": msg.get("internetMessageId", ""),
        "conversation_id":     msg.get("conversationId", ""),
    }


# ── Email send ────────────────────────────────────────────────────────────────

def send_reply(message_id: str, reply_body: str) -> bool:
    """
    Send a threaded reply to a specific email message.
    Returns True on success.
    """
    url = "%s/users/%s/messages/%s/reply" % (GRAPH_BASE, EMAIL_ADDRESS, message_id)
    payload = {"comment": reply_body}
    resp = requests.post(url, headers=_headers(), json=payload, timeout=15)
    if resp.status_code == 202:
        logger.info("Email reply sent for message: %s", message_id)
        return True
    else:
        logger.error("Email send failed %d: %s", resp.status_code, resp.text)
        return False
