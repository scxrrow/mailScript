"""Send the CV to the HR addresses listed in the Excel file, through Proton Mail Bridge.

Proton Mail Bridge must be running on this Mac (it exposes a local SMTP server).
Safe by default: without --send nothing is sent, it only prints what would be.

Environment variables:
    PROTON_USER      your Proton address (used as sender, and as the Bridge username
                     unless BRIDGE_USER is set)
    BRIDGE_PASSWORD  the password shown by Bridge for your account (not your Proton password)
    BRIDGE_USER      optional, the username shown by Bridge if it differs from PROTON_USER
    SMTP_HOST        optional, default 127.0.0.1
    SMTP_PORT        optional, default 1025 (the SMTP port shown by Bridge)
    REPLY_TO         optional, address that replies should go to

Examples:
    uv run python mailScript.py                  # preview (dry run)
    uv run python mailScript.py --test           # send 3 mails to yourself
    uv run python mailScript.py --send --limit 20
"""

import argparse
import csv
import mimetypes
import os
import random
import re
import smtplib
import ssl
import sys
import time
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent
EXCEL = BASE / "Entreprises_Rennes_30km_contacts_RH_SI_v8.xlsx"
CV = BASE / "CV Noah Felten - Alternance Bachelor Cybersécurité.pdf"
TEMPLATE = BASE / "message.txt"
SENT_LOG = BASE / "sent_log.csv"

SHEET = "Entreprises"
COMPANY_COL = "Entreprise"
EMAIL_COL = "Email RH / recrutement"

SMTP_HOST = os.environ.get("SMTP_HOST", "127.0.0.1")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "1025"))

EMAIL_RE = re.compile(r"^[\w.+'-]+@[\w-]+(\.[\w-]+)+$")
PLACEHOLDER_RE = re.compile(r"\[À COMPLÉTER[^\]]*\]")


def pretty_company(name: str) -> str:
    """'KEOLIS RENNES' -> 'Keolis Rennes' (only when the name is all caps)."""
    name = name.strip()
    return name.title() if name.isupper() else name


def load_template() -> tuple[str, str]:
    header, _, body = TEMPLATE.read_text(encoding="utf-8").partition("\n---\n")
    subject = header.removeprefix("Objet:").strip()
    return subject, body.strip() + "\n"


def load_recipients() -> list[tuple[str, str]]:
    """Return unique (email, company) pairs from the HR column."""
    wb = openpyxl.load_workbook(EXCEL, read_only=True, data_only=True)
    rows = wb[SHEET].iter_rows(values_only=True)
    header = next(rows)
    company_i, email_i = header.index(COMPANY_COL), header.index(EMAIL_COL)

    seen: set[str] = set()
    out: list[tuple[str, str]] = []
    for row in rows:
        raw, company = row[email_i], row[company_i]
        if not raw or not company:
            continue
        for addr in re.split(r"[;,\s]+", str(raw).strip()):
            addr = addr.lower()
            if EMAIL_RE.match(addr) and addr not in seen:
                seen.add(addr)
                out.append((addr, pretty_company(str(company))))
    return out


def already_sent() -> set[str]:
    if not SENT_LOG.exists():
        return set()
    with SENT_LOG.open(newline="", encoding="utf-8") as f:
        return {r["email"] for r in csv.DictReader(f) if r["status"] == "sent"}


def log(email: str, company: str, status: str, detail: str = "") -> None:
    new = not SENT_LOG.exists()
    with SENT_LOG.open("a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["timestamp", "email", "company", "status", "detail"])
        w.writerow([datetime.now().isoformat(timespec="seconds"), email, company, status, detail])


def build_message(sender: str, to: str, company: str, subject: str, body: str) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = to
    msg["Subject"] = subject
    if reply_to := os.environ.get("REPLY_TO"):
        msg["Reply-To"] = reply_to
    msg.set_content(body.replace("{entreprise}", company))
    ctype, _ = mimetypes.guess_type(CV)
    maintype, subtype = (ctype or "application/pdf").split("/")
    msg.add_attachment(CV.read_bytes(), maintype=maintype, subtype=subtype, filename=CV.name)
    return msg


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--send", action="store_true", help="really send the mails to the companies")
    p.add_argument("--test", action="store_true", help="send the first mails to yourself instead")
    p.add_argument("--limit", type=int, default=None, help="max number of mails (test default: 3)")
    p.add_argument("--min-delay", type=float, default=20, help="min seconds between mails")
    p.add_argument("--max-delay", type=float, default=45, help="max seconds between mails")
    args = p.parse_args()

    subject, body = load_template()
    recipients = load_recipients()
    done = already_sent()
    todo = [(e, c) for e, c in recipients if e not in done]
    skipped = len(recipients) - len(todo)
    limit = args.limit if args.limit is not None else (3 if args.test else None)
    if limit is not None:
        todo = todo[:limit]

    print(f"{len(recipients)} HR addresses, {skipped} already sent, {len(todo)} to process.")

    if not (args.send or args.test):
        for email, company in todo:
            print(f"[dry-run] {email:45s} {company}")
        if todo:
            print("\n--- first message ---")
            print(f"Objet: {subject}\n\n{body.replace('{entreprise}', todo[0][1])}")
        print("Nothing sent. Use --test, then --send.")
        return 0

    if args.send and PLACEHOLDER_RE.search(subject + body):
        print("Refusing to send: message.txt still contains [À COMPLÉTER ...].", file=sys.stderr)
        return 1

    user, password = os.environ.get("PROTON_USER"), os.environ.get("BRIDGE_PASSWORD")
    if not user or not password:
        print("Set PROTON_USER and BRIDGE_PASSWORD first.", file=sys.stderr)
        return 1
    login = os.environ.get("BRIDGE_USER", user)

    # Bridge uses a self-signed certificate, so verification is only skipped for localhost.
    ctx = ssl.create_default_context()
    if SMTP_HOST in ("127.0.0.1", "localhost"):
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE

    sent = failed = 0
    try:
        smtp = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=60)
    except OSError as exc:
        print(f"Cannot reach Bridge at {SMTP_HOST}:{SMTP_PORT} ({exc}). Is Proton Mail Bridge running?", file=sys.stderr)
        return 1
    with smtp:
        smtp.starttls(context=ctx)
        try:
            smtp.login(login, password)
        except smtplib.SMTPAuthenticationError as exc:
            print(f"Bridge refused the login ({exc.smtp_code}). Check the username and password "
                  "shown in Bridge (Settings > your account > Configure email client).", file=sys.stderr)
            return 1
        for i, (email, company) in enumerate(todo):
            target = user if args.test else email
            msg = build_message(user, target, company, subject, body)
            try:
                smtp.send_message(msg)
            except smtplib.SMTPRecipientsRefused as exc:
                failed += 1
                log(email, company, "refused", str(exc))
                print(f"refused  {email}")
            except (smtplib.SMTPSenderRefused, smtplib.SMTPDataError, smtplib.SMTPServerDisconnected) as exc:
                log(email, company, "error", str(exc))
                print(f"stopping, server error on {email}: {exc}", file=sys.stderr)
                return 1
            else:
                sent += 1
                if not args.test:
                    log(email, company, "sent")
                print(f"sent     {target}  ({company})")
            if i < len(todo) - 1:
                time.sleep(random.uniform(args.min_delay, args.max_delay))
    print(f"Done: {sent} sent, {failed} refused.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
