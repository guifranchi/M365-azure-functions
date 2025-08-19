import os
import json
import re
import logging
import azure.functions as func
import requests
from msal import ConfidentialClientApplication
from bs4 import BeautifulSoup

# === Environment (set these in Azure → Function App → Configuration) ===
TENANT_ID = os.environ.get("TENANT_ID")
CLIENT_ID = os.environ.get("CLIENT_ID")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")
GRAPH_SCOPE = os.environ.get("GRAPH_SCOPE", "https://graph.microsoft.com/.default")
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# === CPF patterns (formatted and unformatted) ===
CPF_FMT = re.compile(r"\b(\d{3})\.(\d{3})\.(\d{3})-(\d{2})\b")
CPF_UNF = re.compile(r"(?<!\d)(\d{11})(?!\d)")

def cpf_checksum_valid(digits: str) -> bool:
    """
    Validates Brazilian CPF checksum.
    - Rejects length != 11
    - Rejects repeated digits (e.g., 00000000000)
    - Validates 2 check digits
    """
    if len(digits) != 11:
        return False
    if digits == digits[0] * 11:
        return False

    # First check digit
    s1 = sum(int(d) * w for d, w in zip(digits[:9], range(10, 1, -1)))
    r1 = (s1 * 10) % 11
    if r1 == 10:
        r1 = 0
    if r1 != int(digits[9]):
        return False

    # Second check digit
    s2 = sum(int(d) * w for d, w in zip(digits[:10], range(11, 1, -1)))
    r2 = (s2 * 10) % 11
    if r2 == 10:
        r2 = 0
    return r2 == int(digits[10])

def get_app_token() -> str:
    """
    App-only token for Microsoft Graph via client credentials.
    Requires Graph application permissions + admin consent.
    """
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_silent([GRAPH_SCOPE], account=None)
    if not result:
        result = app.acquire_token_for_client(scopes=[GRAPH_SCOPE])
    if "access_token" not in result:
        raise RuntimeError(f"Failed to get token: {result}")
    return result["access_token"]

def graph_get(url: str, token: str, headers: dict | None = None) -> dict:
    h = {"Authorization": f"Bearer {token}"}
    if headers:
        h.update(headers)
    resp = requests.get(url, headers=h)
    if resp.status_code >= 400:
        raise RuntimeError(f"Graph GET {url} -> {resp.status_code} {resp.text}")
    return resp.json()

def graph_post(url: str, token: str, payload: dict, headers: dict | None = None) -> dict:
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    if headers:
        h.update(headers)
    resp = requests.post(url, headers=h, data=json.dumps(payload))
    if resp.status_code >= 300:
        raise RuntimeError(f"Graph POST {url} -> {resp.status_code} {resp.text}")
    return resp.json() if resp.text else {}

def html_to_text(html: str) -> str:
    """Render HTML to text (fallback redaction path)."""
    soup = BeautifulSoup(html or "", "lxml")
    return soup.get_text(" ", strip=False)

def find_cpfs_in_html(html: str) -> list[str]:
    """Return original CPF substrings to replace in HTML."""
    originals = []
    for m in CPF_FMT.finditer(html):
        digits = "".join(m.groups())
        if cpf_checksum_valid(digits):
            originals.append(m.group(0))
    for m in CPF_UNF.finditer(html):
        digits = m.group(1)
        if cpf_checksum_valid(digits):
            originals.append(m.group(0))
    return originals

def replace_in_html(html: str, originals: list[str]) -> str:
    """Simple literal replacement in HTML with [REDACTED]."""
    for s in originals:
        html = html.replace(s, "[REDACTED]")
    return html

def redact_text(text: str) -> tuple[str, int]:
    """
    Redact CPFs in plain text (formatted + unformatted).
    Returns (redacted_text, count).
    """
    count = 0

    def repl_fmt(m):
        nonlocal count
        digits = "".join(m.groups())
        if cpf_checksum_valid(digits):
            count += 1
            return "[REDACTED]"
        return m.group(0)

    def repl_unf(m):
        nonlocal count
        digits = m.group(1)
        if cpf_checksum_valid(digits):
            count += 1
            return "[REDACTED]"
        return m.group(0)

    t = CPF_FMT.sub(repl_fmt, text)
    t = CPF_UNF.sub(repl_unf, t)
    return t, count

def build_sendmail_payload(original_msg: dict, new_body_html: str) -> dict:
    return {
        "message": {
            "subject": original_msg.get("subject", "(no subject)"),
            "body": { "contentType": "HTML", "content": new_body_html },
            "toRecipients": original_msg.get("toRecipients", []),
            "ccRecipients": original_msg.get("ccRecipients", []),
            "bccRecipients": original_msg.get("bccRecipients", []),
            "replyTo": original_msg.get("replyTo", [])
        },
        "saveToSentItems": True
    }

async def main(req: func.HttpRequest) -> func.HttpResponse:
    """
    Input JSON:
    {
      "userPrincipalName": "sender@contoso.com",
      "graphMessageId": "AAMkAD...",
      "internetMessageId": "<abc@contoso.com>"
    }
    """
    try:
        body = req.get_json()
    except Exception:
        return func.HttpResponse(
            json.dumps({"error": "invalid_json"}),
            status_code=400, mimetype="application/json"
        )

    user_upn = body.get("userPrincipalName")
    graph_message_id = body.get("graphMessageId")
    internet_message_id = body.get("internetMessageId")

    if not user_upn or not (graph_message_id or internet_message_id):
        return func.HttpResponse(
            json.dumps({"error": "missing userPrincipalName and message id"}),
            status_code=400, mimetype="application/json"
        )

    try:
        token = get_app_token()

        # Resolve the message via Graph
        if graph_message_id:
            msg = graph_get(
                f"{GRAPH_BASE}/users/{user_upn}/messages/{graph_message_id}"
                "?$select=subject,body,toRecipients,ccRecipients,bccRecipients,replyTo",
                token
            )
        else:
            # internetMessageId search (must include angle brackets)
            url = (f"{GRAPH_BASE}/users/{user_upn}/messages"
                   f"?$filter=internetMessageId eq '{internet_message_id}'&$top=1")
            msg_search = graph_get(url, token, headers={"ConsistencyLevel": "eventual"})
            values = msg_search.get("value", [])
            if not values:
                raise RuntimeError("message_not_found_by_internetMessageId")
            msg = values[0]

        body_html = (msg.get("body") or {}).get("content", "") or ""

        # Primary: redact directly in HTML by replacing exact substrings
        originals = find_cpfs_in_html(body_html)
        new_body_html = replace_in_html(body_html, originals)

        # Fallback: if none found in HTML, try text rendering & replace whole body with <pre> text
        if not originals:
            text = html_to_text(body_html)
            red_text, count = redact_text(text)
            if count > 0:
                new_body_html = f"<pre>{red_text}</pre>"

        # Send sanitized mail
        payload = build_sendmail_payload(msg, new_body_html)
        graph_post(f"{GRAPH_BASE}/users/{user_upn}/sendMail", token, payload)

        return func.HttpResponse(
            json.dumps({"status": "resent", "messageId": msg.get("id")}),
            status_code=200, mimetype="application/json"
        )

    except Exception as e:
        logging.exception("processing_failed")
        return func.HttpResponse(
            json.dumps({"status": "error", "detail": str(e)}),
            status_code=500, mimetype="application/json"
        )
