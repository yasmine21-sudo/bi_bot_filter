from __future__ import annotations

import json
import os
from pathlib import Path
import pytest
from unittest.mock import MagicMock

import main
from main import (
    RecipientGroup,
    Settings,
    _col_ref_to_index,
    load_recipient_mappings,
    normalize_option_key,
    parse_email_list,
    sanitize_filename,
)
import webapp


def test_col_ref_to_index():
    assert _col_ref_to_index("A") == 0
    assert _col_ref_to_index("B") == 1
    assert _col_ref_to_index("Z") == 25
    assert _col_ref_to_index("AA") == 26
    assert _col_ref_to_index("AB") == 27
    assert _col_ref_to_index("A1") == 0
    assert _col_ref_to_index("C12") == 2


def test_normalize_and_sanitize():
    assert normalize_option_key("  GAMMA   SECURITY  ") == "gamma security"
    assert sanitize_filename("BTPH (Direction) & Co!") == "BTPH_Direction_Co"


def test_parse_email_list():
    assert parse_email_list("a@x.com; b@y.com, c@z.com") == ["a@x.com", "b@y.com", "c@z.com"]
    assert parse_email_list("") == []


def test_load_recipient_mappings(tmp_path):
    xlsx_path = tmp_path / "test_destinataire.xlsx"
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Filiale", "Email distinataire AA", "Email cc"])
    ws.append(["Filiale A", "user1@a.com", "cc1@a.com"])
    ws.append(["Filiale A", "user2@a.com", "cc1@a.com; cc2@a.com"])
    wb.save(xlsx_path)

    mappings = load_recipient_mappings(str(xlsx_path))
    assert "filiale a" in mappings
    group = mappings["filiale a"]
    assert group.to == ["user1@a.com", "user2@a.com"]
    assert group.cc == ["cc1@a.com", "cc2@a.com"]


def test_run_full_override_safety(monkeypatch):
    monkeypatch.setenv("TEST_EMAIL_TO", "test_override@example.com")
    import importlib
    import run_full_override
    importlib.reload(run_full_override)

    mock_settings = Settings(
        report_url="http://localhost",
        output_dir=Path("output"),
        log_dir=Path("logs"),
        headless=True,
        browser_channel=None,
        viewport_width=1920,
        viewport_height=1080,
        device_scale_factor=2.0,
        navigation_timeout_ms=30000,
        report_render_timeout_ms=30000,
        report_stable_interval_ms=1000,
        report_stable_polls=2,
        post_tab_click_wait_ms=1000,
        timezone="UTC",
        auth_mode="none",
        auth_server_whitelist="",
        edge_user_data_dir=None,
        edge_profile_directory=None,
        http_username=None,
        http_password=None,
        pbirs_username=None,
        pbirs_password=None,
        login_username_selector=None,
        login_password_selector=None,
        login_submit_selector=None,
        smtp_host="localhost",
        smtp_port=25,
        smtp_timeout_seconds=30,
        smtp_use_tls=False,
        smtp_use_ssl=False,
        smtp_skip_verify=True,
        smtp_username=None,
        smtp_password=None,
        smtp_envelope_from=None,
        email_from="bot@example.com",
        email_reply_to=None,
        email_to=["production@realcompany.com"],
        expected_sheets=[],
        filter_slicer_name=None,
        filter_slicer_page=None,
        filter_exclude_options=[],
        max_workers=1,
        slicer_dropdown_wait_ms=0,
        slicer_apply_wait_ms=0,
    )

    monkeypatch.setattr(run_full_override, "original_load_settings", lambda *a, **kw: mock_settings)
    s = main.load_settings()
    assert s.email_to == ["test_override@example.com"]

    monkeypatch.setattr(
        run_full_override,
        "original_load_recipient_mappings",
        lambda *a, **kw: {"filiale a": RecipientGroup(to=["real@company.com"], cc=["cc@company.com"])},
    )
    m = main.load_recipient_mappings()
    assert m["filiale a"].to == ["test_override@example.com"]
    assert m["filiale a"].cc == []


def test_slicer_escaping_locators():
    mock_frame = MagicMock()
    mock_container = MagicMock()
    mock_dropdown = MagicMock()
    mock_frame.locator.return_value.filter.return_value = mock_container
    mock_container.locator.return_value = mock_dropdown

    slicer_name = "BTPH (Direction) & O'Reilly"
    res = mock_frame.locator(".slicer-container").filter(has_text=slicer_name).locator(".slicer-dropdown-menu")

    mock_frame.locator.assert_called_with(".slicer-container")
    mock_frame.locator.return_value.filter.assert_called_with(has_text=slicer_name)


def test_webapp_api():
    client = webapp.app.test_client()

    res_rec = client.get("/api/recipients")
    assert res_rec.status_code == 200

    res_set = client.get("/api/settings")
    assert res_set.status_code == 200

    res_stat = client.get("/api/status")
    assert res_stat.status_code == 200

    res_stop = client.post("/api/stop")
    assert res_stop.status_code == 400  # no run process currently running
