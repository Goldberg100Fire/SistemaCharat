import base64
from email.message import EmailMessage
import os

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build


def _get_gmail_service() -> tuple[object, str]:
    """
    Returns (gmail_service, sender_email).
    Requires env vars:
      - GMAIL_CLIENT_ID
      - GMAIL_CLIENT_SECRET
      - GMAIL_REFRESH_TOKEN
      - GMAIL_SENDER_EMAIL
    """
    client_id = os.environ.get("GMAIL_CLIENT_ID")
    client_secret = os.environ.get("GMAIL_CLIENT_SECRET")
    refresh_token = os.environ.get("GMAIL_REFRESH_TOKEN")
    sender = os.environ.get("GMAIL_SENDER_EMAIL")

    missing = [k for k, v in {
        "GMAIL_CLIENT_ID": client_id,
        "GMAIL_CLIENT_SECRET": client_secret,
        "GMAIL_REFRESH_TOKEN": refresh_token,
        "GMAIL_SENDER_EMAIL": sender,
    }.items() if not v]
    if missing:
        raise RuntimeError(f"Faltan variables de entorno para Gmail API: {', '.join(missing)}")

    creds = Credentials(
        None,
        refresh_token=refresh_token,
        token_uri="https://oauth2.googleapis.com/token",
        client_id=client_id,
        client_secret=client_secret,
        scopes=["https://www.googleapis.com/auth/gmail.send"],
    )

    service = build("gmail", "v1", credentials=creds, cache_discovery=False)
    return service, sender


def send_email(to_email: str, subject: str, body_text: str) -> None:
    service, sender = _get_gmail_service()

    msg = EmailMessage()
    msg["To"] = to_email
    msg["From"] = sender
    msg["Subject"] = subject
    msg.set_content(body_text)

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
    service.users().messages().send(userId="me", body={"raw": raw}).execute()

