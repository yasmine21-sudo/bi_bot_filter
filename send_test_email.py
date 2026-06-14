from __future__ import annotations

import logging
import smtplib
import sys
from email.message import EmailMessage

from main import (
    build_logger,
    ensure_directories,
    load_settings,
    validate_settings,
    validate_timezone,
)


def send_test_email() -> int:
    settings = load_settings()
    ensure_directories(settings)
    logger = build_logger(settings)
    validate_settings(settings)
    validate_timezone(settings, logger)

    envelope_from = settings.smtp_envelope_from or settings.email_from
    subject = "SMTP Validation Success"
    body = "\n".join(
        [
            "SMTP validation message.",
            "If you received this email, the sender identity and SMTP submission are working.",
            f"Authenticated SMTP user: {settings.smtp_username or '(anonymous)'}",
            f"Header From: {settings.email_from}",
            f"Envelope From: {envelope_from}",
            f"Recipients: {', '.join(settings.email_to)}",
        ]
    )

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = settings.email_from
    message["To"] = ", ".join(settings.email_to)
    if settings.email_reply_to:
        message["Reply-To"] = settings.email_reply_to
    message.set_content(body)

    logger.info("Sending SMTP validation email to %s", settings.email_to)
    with smtplib.SMTP(settings.smtp_host, settings.smtp_port, timeout=60) as smtp:
        smtp.ehlo()
        if settings.smtp_use_tls:
            smtp.starttls()
            smtp.ehlo()
        if settings.smtp_username and settings.smtp_password:
            smtp.login(settings.smtp_username, settings.smtp_password)
        smtp.send_message(message, from_addr=envelope_from)

    logger.info("SMTP validation email sent successfully.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(send_test_email())
    except Exception as error:
        logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
        logging.exception("SMTP validation failed: %s", error)
        raise SystemExit(1)
