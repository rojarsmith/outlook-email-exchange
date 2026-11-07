# Outlook Email Exchange

Listens for new mail in the Outlook desktop client on Windows and forwards each message (subject, plain-text and HTML body) to another mailbox.

```
Outlook (OnNewMailEx) ──► filter ──► SMTP / Mailjet / Mailgun / Courier Hub ──► recipient
```

## Features

- Real-time trigger through Outlook COM events (`win32com`); no IMAP or Graph API needed.
- Four delivery back ends, selected with `exchange.mode`:

  | Mode | How it sends | Notes |
  |---|---|---|
  | `smtp` | `smtplib` | Port 465 uses implicit SSL, other ports use STARTTLS. Works with GoDaddy, Gmail, Microsoft 365, etc. |
  | `mailjet` | Mailjet Send API v3.1 | `sender.email` must be a verified Mailjet sender. |
  | `mailgun` | Mailgun REST API | Sends from `mailgun@<mailgun.domain>`. |
  | `courierhub` | Courier Hub REST API `POST /v1/emails` | Bearer token auth; an `Idempotency-Key` derived from the Outlook item ID prevents duplicate sends. |

- Forwards HTML bodies when present, with plain text as a fallback.
- Keyword filter: mail whose body contains any keyword in `filter.keywords` is skipped.
- One failing message never stops the listener; errors are printed and processing continues.

## Requirements

- Windows with **Outlook desktop** installed, signed in, and running (classic Outlook; the "new Outlook" has no COM interface).
- Python 3.12 and [pipenv](https://pipenv.pypa.io/).

## Installation

```shell
git clone https://github.com/rojarsmith/outlook-email-exchange.git
cd outlook-email-exchange
pipenv install
copy config.properties.template config.properties
```

Then edit `config.properties` (it is listed in `.gitignore`, so your credentials stay out of git).

## Configuration

| Key | Required for | Description |
|---|---|---|
| `exchange.mode` | all | `smtp`, `mailjet`, `mailgun` or `courierhub` |
| `recipient` | all | Address that receives the forwarded mail |
| `sender.name` | optional | Display name of the sender |
| `sender.email` | optional | Sender address; defaults to `smtp.user` |
| `filter.keywords` | optional | Comma-separated; bodies containing any keyword are not forwarded |
| `smtp.server` / `smtp.port` / `smtp.user` / `smtp.password` | `smtp` | SMTP account |
| `smtp.security` | `smtp` | `auto` (default), `ssl` or `starttls` |
| `mailjet.api.key` / `mailjet.api.secret` | `mailjet` | Mailjet API credentials |
| `mailgun.domain` / `mailgun.api` | `mailgun` | Mailgun sending domain and API key |
| `courierhub.url` / `courierhub.api.key` | `courierhub` | Courier Hub base URL (e.g. `https://courierhub.example.io`) and API key |

### Example: GoDaddy email (SMTP)

GoDaddy **Professional / Workspace Email** (the `secureserver.net` mailboxes):

```properties
exchange.mode=smtp
smtp.server=smtpout.secureserver.net
smtp.port=465
smtp.user=you@yourdomain.com
smtp.password=<mailbox password>
sender.name=Your Name
recipient=you@gmail.com
```

| Setting | Value |
|---|---|
| Outgoing server | `smtpout.secureserver.net` |
| Port | `465` with implicit SSL (recommended), or `587` with STARTTLS |
| Username | The **full** email address |
| Password | The mailbox password (not the GoDaddy account password) |
| Regional alternatives | `smtpout.asia.secureserver.net`, `smtpout.europe.secureserver.net` |

Notes:

- `smtp.security` picks the encryption: `auto` (default: `ssl` on port 465, `starttls` on other ports), `ssl` (`smtplib.SMTP_SSL`), or `starttls` (plain connection, then `STARTTLS`).
- The original GoDaddy script used port `465` with `starttls`. If that is what works on your network, set `smtp.security=starttls` to reproduce it exactly. If the connection times out, use the default `auto` (SSL on 465).
- If your GoDaddy mail is **Microsoft 365 from GoDaddy**, use `smtp.office365.com`, port `587`, and make sure *Authenticated SMTP* is enabled for the mailbox in the Microsoft 365 admin center.
- `sender.email` (or `smtp.user`) must be the GoDaddy mailbox itself, otherwise the server rejects the message.

### Example: Courier Hub

```properties
exchange.mode=courierhub
courierhub.url=https://courierhub.example.io
courierhub.api.key=<api key>
recipient=you@gmail.com
```

A successful request returns HTTP `202`, and the script prints the queued job ID and status.

## Usage

Start Outlook first, then run:

```shell
pipenv run python email-exchange.py
```

or double-click `email-exchange.bat`. It switches to the script's own folder, so it works wherever the repository is cloned.

Stop it with `Ctrl+C`.

### Run at sign-in (optional)

Put a shortcut to `email-exchange.bat` in `shell:startup`, or create a Task Scheduler task with the trigger *At log on* and the action set to `email-exchange.bat`. Outlook must also start at sign-in.

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `exchange.mode must be one of ...` | `exchange.mode` is empty or misspelled in `config.properties`. |
| `FileNotFoundError: config.properties` | Run from the repository folder, or use the `.bat`. |
| Nothing happens when mail arrives | Outlook must be running in the same user session; the "new Outlook" app is not supported. Check whether the body matches `filter.keywords`. |
| `SMTPAuthenticationError` | Use the full email address as the username; for Microsoft 365, enable Authenticated SMTP. |
| `SMTPServerDisconnected` / timeout | The encryption doesn't match the server. Switch `smtp.security` between `ssl` and `starttls`, or use port 587 with `starttls`. |
| Courier Hub `401` / `403` | Wrong `courierhub.api.key`. |
| Mailjet `401` / sender error | API key wrong, or `sender.email` not verified in Mailjet. |
