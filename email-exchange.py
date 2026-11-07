import smtplib
import time
import uuid
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr

import pythoncom
import requests
import win32com.client
from jproperties import Properties

# ==========================================
# Configuration (config.properties)
# ==========================================

with open("config.properties", "r+b") as f:
    p = Properties()
    p.load(f, "utf-8")


def cfg(key, default=""):
    value = p.properties.get(key)
    return value.strip() if value and value.strip() else default


exchange_mode = cfg("exchange.mode").lower()

smtp_server = cfg("smtp.server")
smtp_port = int(cfg("smtp.port", "587"))
smtp_user = cfg("smtp.user")
smtp_password = cfg("smtp.password")
# auto (465 -> ssl, other ports -> starttls) | ssl | starttls
smtp_security = cfg("smtp.security", "auto").lower()
if smtp_security == "auto":
    smtp_security = "ssl" if smtp_port == 465 else "starttls"

mailjet_api_key = cfg("mailjet.api.key")
mailjet_api_secret = cfg("mailjet.api.secret")

mailgun_domain = cfg("mailgun.domain")
mailgun_api = cfg("mailgun.api")

courierhub_url = cfg("courierhub.url")
courierhub_api_key = cfg("courierhub.api.key")

sender_email = cfg("sender.email", smtp_user)
sender_name = cfg("sender.name", "Outlook Email Exchange")
recipient = cfg("recipient")

# Mail whose body contains any of these keywords is NOT forwarded.
filter_keywords = [k.strip() for k in cfg("filter.keywords").split(",") if k.strip()]

SUPPORTED_MODES = ("smtp", "mailjet", "mailgun", "courierhub")


# ==========================================
# Outlook
# ==========================================


class OutlookHandler:
    def __init__(self):
        self.outlook = win32com.client.DispatchWithEvents(
            "Outlook.Application", OutlookEvents
        )


class OutlookEvents:
    def OnNewMailEx(self, received_items_ids):
        for mail_id in received_items_ids.split(","):
            try:
                mail = win32com.client.Dispatch(
                    "Outlook.Application"
                ).Session.GetItemFromID(mail_id)

                subject = mail.Subject or "(No Subject)"
                body = mail.Body or ""
                html_body = getattr(mail, "HTMLBody", "") or ""

                print("\n================================")
                print("Receive:", subject)
                print("================================")

                matched = [k for k in filter_keywords if k in body]
                if matched:
                    print("Filtered by", matched, "- skip.")
                    continue

                send_notification(subject, body, html_body, mail_id)
            except Exception as e:
                print("Outlook processing error:", e)


# ==========================================
# Senders
# ==========================================


def send_smtp(subject, body, html_body):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = formataddr((sender_name, sender_email))
    msg["To"] = recipient
    msg.attach(MIMEText(body, "plain", "utf-8"))
    if html_body.strip():
        msg.attach(MIMEText(html_body, "html", "utf-8"))

    if smtp_security == "ssl":
        server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=30)
    else:
        server = smtplib.SMTP(smtp_server, smtp_port, timeout=30)
        server.ehlo()
        server.starttls()
        server.ehlo()
    try:
        server.login(smtp_user, smtp_password)
        refused = server.sendmail(sender_email, [recipient], msg.as_string())
        print("<SUCCESS> SMTP" if not refused else f"<FAILED> SMTP {refused}")
    finally:
        server.quit()


def send_mailjet(subject, body, html_body):
    from mailjet_rest import Client

    mailjet = Client(auth=(mailjet_api_key, mailjet_api_secret), version="v3.1")
    message = {
        "From": {"Email": sender_email, "Name": sender_name},
        "To": [{"Email": recipient}],
        "Subject": subject,
        "TextPart": body,
    }
    if html_body.strip():
        message["HTMLPart"] = html_body
    result = mailjet.send.create(data={"Messages": [message]})
    print("Mailjet:", result.status_code, result.json())


def send_mailgun(subject, body, html_body):
    data = {
        "from": f"{sender_name} <mailgun@{mailgun_domain}>",
        "to": [recipient],
        "subject": subject,
        "text": body,
    }
    if html_body.strip():
        data["html"] = html_body
    response = requests.post(
        f"https://api.mailgun.net/v3/{mailgun_domain}/messages",
        auth=("api", mailgun_api),
        data=data,
        timeout=30,
    )
    print("Mailgun:", response.status_code, response.text)


def send_courierhub(subject, body, html_body, mail_id=None):
    url = f"{courierhub_url.rstrip('/')}/v1/emails"
    # Same Outlook mail -> same key, so Courier Hub will not send it twice.
    idempotency_key = (
        uuid.uuid5(uuid.NAMESPACE_URL, str(mail_id)).hex
        if mail_id
        else uuid.uuid4().hex
    )
    headers = {
        "Authorization": f"Bearer {courierhub_api_key}",
        "Content-Type": "application/json",
        "Idempotency-Key": idempotency_key,
    }
    data = {
        "to": [recipient],
        "subject": subject,
        "text": body if body.strip() else "(Empty email body)",
    }
    if html_body.strip():
        data["html"] = html_body

    response = requests.post(url, headers=headers, json=data, timeout=30)
    if response.status_code == 202:
        result = response.json()
        print(
            "<SUCCESS> Courier Hub queued. Job:",
            result.get("id"),
            "Status:",
            result.get("status"),
        )
    else:
        print("<FAILED> Courier Hub", response.status_code, response.text)


def send_notification(subject, body, html_body="", mail_id=None):
    print(f"Forwarding via {exchange_mode} to {recipient} ...")
    try:
        if exchange_mode == "smtp":
            send_smtp(subject, body, html_body)
        elif exchange_mode == "mailjet":
            send_mailjet(subject, body, html_body)
        elif exchange_mode == "mailgun":
            send_mailgun(subject, body, html_body)
        elif exchange_mode == "courierhub":
            send_courierhub(subject, body, html_body, mail_id)
    except requests.exceptions.Timeout:
        print("ERROR: request timeout")
    except requests.exceptions.ConnectionError:
        print("ERROR: cannot connect to the API server")
    except Exception as e:
        print("Send error:", e)


# ==========================================
# Main
# ==========================================

if __name__ == "__main__":
    if exchange_mode not in SUPPORTED_MODES:
        raise SystemExit(
            f"exchange.mode must be one of {SUPPORTED_MODES}, got '{exchange_mode}'"
        )
    if not recipient:
        raise SystemExit("recipient is not set in config.properties")

    print("================================")
    print("Outlook Email Exchange")
    print("Mode:", exchange_mode, "| Recipient:", recipient)
    print("Filters:", filter_keywords or "(none)")
    print("================================")

    outlook_handler = OutlookHandler()
    print("Outlook monitoring started... (Ctrl+C to stop)")

    try:
        while True:
            pythoncom.PumpWaitingMessages()
            time.sleep(0.1)
    except KeyboardInterrupt:
        print("\nService stopped.")
