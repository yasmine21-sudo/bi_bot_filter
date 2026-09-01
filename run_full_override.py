"""Test run: send every filiale's email to a single test address.

The address is read from TEST_EMAIL_TO (set it in the web console's "Mode test"
field, or in .env). This overrides both the To and CC of every recipient group so
nothing reaches the real filiales.
"""
import os
import sys

import main
from main import load_dotenv_file, load_recipient_mappings, load_settings

# Populate os.environ from .env (values already set in the environment win).
load_dotenv_file()
TEST_EMAIL_TO = (os.getenv("TEST_EMAIL_TO") or "").strip()

# Store references to original functions
original_load_recipient_mappings = load_recipient_mappings
original_load_settings = load_settings


def mock_load_settings(*args, **kwargs):
    if not TEST_EMAIL_TO:
        raise ValueError(
            "TEST_EMAIL_TO is not set. Enter a test address in the web console's "
            "'Mode test' field (or set TEST_EMAIL_TO in .env) before running test mode."
        )
    settings = original_load_settings(*args, **kwargs)
    settings.email_to = [TEST_EMAIL_TO]
    return settings


def mock_load_recipient_mappings(*args, **kwargs):
    if not TEST_EMAIL_TO:
        raise ValueError(
            "TEST_EMAIL_TO is not set. Enter a test address in the web console's "
            "'Mode test' field (or set TEST_EMAIL_TO in .env) before running test mode."
        )
    mappings = original_load_recipient_mappings(*args, **kwargs)
    for group in mappings.values():
        group.to = [TEST_EMAIL_TO]
        group.cc = []
    return mappings


# Patch original functions with overrides
main.load_settings = mock_load_settings
main.load_recipient_mappings = mock_load_recipient_mappings

if __name__ == "__main__":
    sys.exit(main.main())

