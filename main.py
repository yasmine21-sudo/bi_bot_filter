from __future__ import annotations

import logging
import os
import re
import smtplib
import ssl
import sys
import time
from urllib.parse import urlparse
from dataclasses import dataclass, field
from datetime import datetime
from datetime import timedelta
from datetime import timezone as dt_timezone
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo
from zoneinfo import ZoneInfoNotFoundError

try:
    from playwright.sync_api import BrowserContext, Page, Playwright, sync_playwright
    from playwright._impl._errors import Error as PlaywrightError
except ImportError:
    BrowserContext = Any
    Page = Any
    Playwright = Any
    sync_playwright = None
    PlaywrightError = Exception


VISUAL_SELECTORS = [
    ".visual-container",
    ".visualContainer",
    "[data-automation-type='visualContainer']",
    "[data-testid*='visual']",
    ".reportCanvas",
    ".reportVisual",
    ".pageContent",
    "canvas",
    "svg",
]

SPINNER_SELECTORS = [
    ".loading",
    ".loadingSpinner",
    ".spinner",
    ".busyIndicator",
    ".loader",
    "[role='progressbar']",
    "[aria-busy='true']",
]

TAB_SELECTORS = [
    "[role='tab']",
    "button[role='tab']",
    "[data-testid*='tab']",
    "[aria-selected]",
    ".pivot-header-tab",
    ".tab-nav-button",
    ".navigationTab",
    ".pageTab",
]

PAGE_THUMBNAIL_SELECTORS = [
    ".section.dynamic.thumbnail-container",
    ".thumbnail-container.section",
    ".section.thumbnail-container",
]

USERNAME_SELECTORS = [
    "input[name='username']",
    "input[name='userName']",
    "input[name='email']",
    "input[type='email']",
    "input[type='text']",
]

PASSWORD_SELECTORS = [
    "input[name='password']",
    "input[type='password']",
]

SUBMIT_SELECTORS = [
    "button[type='submit']",
    "input[type='submit']",
    "button:has-text('Sign in')",
    "button:has-text('Log in')",
    "button:has-text('Login')",
]


@dataclass
class Settings:
    report_url: str
    output_dir: Path
    log_dir: Path
    headless: bool
    browser_channel: str | None
    viewport_width: int
    viewport_height: int
    device_scale_factor: float
    navigation_timeout_ms: int
    report_render_timeout_ms: int
    report_stable_interval_ms: int
    report_stable_polls: int
    post_tab_click_wait_ms: int
    screenshot_prefix: str
    timezone: str
    auth_mode: str
    auth_server_whitelist: str
    edge_user_data_dir: str | None
    edge_profile_directory: str | None
    http_username: str | None
    http_password: str | None
    pbirs_username: str | None
    pbirs_password: str | None
    login_username_selector: str | None
    login_password_selector: str | None
    login_submit_selector: str | None
    smtp_host: str
    smtp_port: int
    smtp_use_tls: bool
    smtp_use_ssl: bool
    smtp_skip_verify: bool
    smtp_username: str | None
    smtp_password: str | None
    smtp_envelope_from: str | None
    email_from: str
    email_reply_to: str | None
    email_to: list[str]
    email_subject_prefix: str
    expected_sheets: list[str]
    filter_slicer_name: str | None
    filter_slicer_page: str | None
    filter_exclude_options: list[str]

    @property
    def browser_profile_dir(self) -> Path:
        return Path(".browser-profile")


@dataclass
class ReportTab:
    label: str
    mode: str
    page_name: str | None = None
    dom_index: int | None = None
    is_active: bool = False


def parse_bool(value: str | None, default: bool = False) -> bool:
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


def get_env(name: str, default: str | None = None) -> str | None:
    value = os.getenv(name)
    if value is None or value == "":
        return default
    return value


def parse_csv(value: str | None) -> list[str]:
    if not value:
        return []
    return [item.strip() for item in value.split(",") if item.strip()]


def load_dotenv_file(path: str = ".env") -> None:
    env_path = Path(path)
    if not env_path.exists():
        return

    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        os.environ.setdefault(key, value)


def load_settings() -> Settings:
    load_dotenv_file()
    report_url = get_env("REPORT_URL")
    if not report_url:
        raise ValueError("REPORT_URL is required.")

    email_to = parse_csv(get_env("EMAIL_TO"))
    if not email_to:
        raise ValueError("EMAIL_TO must contain at least one recipient.")

    smtp_port_val = int(get_env("SMTP_PORT", "587"))
    smtp_use_ssl_val = get_env("SMTP_USE_SSL")
    if smtp_use_ssl_val is None:
        smtp_use_ssl = (smtp_port_val == 465)
    else:
        smtp_use_ssl = parse_bool(smtp_use_ssl_val, default=False)

    smtp_use_tls_val = get_env("SMTP_USE_TLS")
    if smtp_use_tls_val is None:
        smtp_use_tls = not smtp_use_ssl
    else:
        smtp_use_tls = parse_bool(smtp_use_tls_val, default=True)

    smtp_skip_verify = parse_bool(get_env("SMTP_SKIP_VERIFY"), default=False)

    return Settings(
        report_url=report_url,
        output_dir=Path(get_env("OUTPUT_DIR", "output")),
        log_dir=Path(get_env("LOG_DIR", "logs")),
        headless=parse_bool(get_env("HEADLESS"), default=True),
        browser_channel=get_env("BROWSER_CHANNEL"),
        viewport_width=int(get_env("VIEWPORT_WIDTH", "1920")),
        viewport_height=int(get_env("VIEWPORT_HEIGHT", "1080")),
        device_scale_factor=float(get_env("DEVICE_SCALE_FACTOR", "2")),
        navigation_timeout_ms=int(get_env("NAVIGATION_TIMEOUT_MS", "120000")),
        report_render_timeout_ms=int(get_env("REPORT_RENDER_TIMEOUT_MS", "180000")),
        report_stable_interval_ms=int(get_env("REPORT_STABLE_INTERVAL_MS", "2000")),
        report_stable_polls=int(get_env("REPORT_STABLE_POLLS", "3")),
        post_tab_click_wait_ms=int(get_env("POST_TAB_CLICK_WAIT_MS", "5000")),
        screenshot_prefix=get_env("SCREENSHOT_PREFIX", "pbirs") or "pbirs",
        timezone=get_env("TIMEZONE", "UTC") or "UTC",
        auth_mode=(get_env("AUTH_MODE", "none") or "none").strip().lower(),
        auth_server_whitelist=get_env("AUTH_SERVER_WHITELIST", "") or "",
        edge_user_data_dir=get_env("EDGE_USER_DATA_DIR"),
        edge_profile_directory=get_env("EDGE_PROFILE_DIRECTORY"),
        http_username=get_env("HTTP_USERNAME"),
        http_password=get_env("HTTP_PASSWORD"),
        pbirs_username=get_env("PBIRS_USERNAME"),
        pbirs_password=get_env("PBIRS_PASSWORD"),
        login_username_selector=get_env("LOGIN_USERNAME_SELECTOR"),
        login_password_selector=get_env("LOGIN_PASSWORD_SELECTOR"),
        login_submit_selector=get_env("LOGIN_SUBMIT_SELECTOR"),
        smtp_host=get_env("SMTP_HOST", "") or "",
        smtp_port=smtp_port_val,
        smtp_use_tls=smtp_use_tls,
        smtp_use_ssl=smtp_use_ssl,
        smtp_skip_verify=smtp_skip_verify,
        smtp_username=get_env("SMTP_USERNAME"),
        smtp_password=get_env("SMTP_PASSWORD"),
        smtp_envelope_from=get_env("SMTP_ENVELOPE_FROM"),
        email_from=get_env("EMAIL_FROM", "") or "",
        email_reply_to=get_env("EMAIL_REPLY_TO"),
        email_to=email_to,
        email_subject_prefix=get_env("EMAIL_SUBJECT_PREFIX", "PBIRS Daily Capture") or "PBIRS Daily Capture",
        expected_sheets=parse_csv(get_env("EXPECTED_SHEETS")),
        filter_slicer_name=get_env("FILTER_SLICER_NAME"),
        filter_slicer_page=get_env("FILTER_SLICER_PAGE"),
        filter_exclude_options=parse_csv(get_env("FILTER_EXCLUDE_OPTIONS")),
    )


def ensure_directories(settings: Settings) -> None:
    settings.output_dir.mkdir(parents=True, exist_ok=True)
    settings.log_dir.mkdir(parents=True, exist_ok=True)
    settings.browser_profile_dir.mkdir(parents=True, exist_ok=True)


def build_logger(settings: Settings) -> logging.Logger:
    logger = logging.getLogger("pbirs_capture")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    file_handler = logging.FileHandler(settings.log_dir / "pbirs_capture.log", encoding="utf-8")
    file_handler.setFormatter(formatter)

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)
    return logger


def now_in_tz(timezone_name: str) -> datetime:
    try:
        return datetime.now(ZoneInfo(timezone_name))
    except ZoneInfoNotFoundError:
        offset_match = re.fullmatch(r"([+-])(\d{2}):(\d{2})", timezone_name.strip())
        if offset_match:
            sign, hours_text, minutes_text = offset_match.groups()
            delta = timedelta(hours=int(hours_text), minutes=int(minutes_text))
            if sign == "-":
                delta = -delta
            return datetime.now(dt_timezone(delta))
        return datetime.now()


def timestamp_compact(timezone_name: str) -> str:
    return now_in_tz(timezone_name).strftime("%Y%m%d_%H%M%S")


def timestamp_readable(timezone_name: str) -> str:
    return now_in_tz(timezone_name).strftime("%Y-%m-%d %H:%M:%S %Z")


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())
    return cleaned.strip("._") or "sheet"


def require_email_config(settings: Settings) -> None:
    required = {
        "SMTP_HOST": settings.smtp_host,
        "EMAIL_FROM": settings.email_from,
    }
    missing = [name for name, value in required.items() if not value]
    if missing:
        raise ValueError(f"Missing required email configuration: {', '.join(missing)}")


def validate_settings(settings: Settings) -> None:
    if "@" in settings.smtp_host:
        raise ValueError(
            "SMTP_HOST must be an SMTP server hostname or IP, not an email address. "
            "Example: smtp.company.local or mail.groupe-hasnaoui.com"
        )

    if settings.auth_mode == "basic" and (not settings.http_username or not settings.http_password):
        raise ValueError("AUTH_MODE=basic requires HTTP_USERNAME and HTTP_PASSWORD.")

    if settings.auth_mode == "form" and (not settings.pbirs_username or not settings.pbirs_password):
        raise ValueError("AUTH_MODE=form requires PBIRS_USERNAME and PBIRS_PASSWORD.")

    if settings.auth_mode not in {"none", "basic", "form", "integrated"}:
        raise ValueError("AUTH_MODE must be one of: none, basic, form, integrated.")


def is_ipv4_host(hostname: str | None) -> bool:
    if not hostname:
        return False
    return bool(re.fullmatch(r"(?:\d{1,3}\.){3}\d{1,3}", hostname))


def build_auth_server_allowlist(settings: Settings) -> str:
    report_host = (urlparse(settings.report_url).hostname or "").strip()
    configured = [item.strip() for item in settings.auth_server_whitelist.split(",") if item.strip()]

    if report_host and report_host not in configured:
        configured.append(report_host)

    return ",".join(dict.fromkeys(configured))


def normalize_sheet_name(value: str) -> str:
    return re.sub(r"\s+", " ", value.strip()).casefold()


def validate_timezone(settings: Settings, logger: logging.Logger) -> None:
    try:
        ZoneInfo(settings.timezone)
        return
    except ZoneInfoNotFoundError:
        pass

    if re.fullmatch(r"[+-]\d{2}:\d{2}", settings.timezone.strip()):
        logger.warning(
            "TIMEZONE '%s' is being used as a fixed UTC offset.",
            settings.timezone,
        )
        return

    logger.warning(
        "TIMEZONE '%s' is not available on this Python installation. "
        "Falling back to the server local time. "
        "If you want a stable offset without tzdata, use something like '+01:00'.",
        settings.timezone,
    )


def build_context(playwright: Playwright, settings: Settings) -> tuple[BrowserContext, Page]:
    common_context_args: dict[str, Any] = {
        "ignore_https_errors": True,
        "viewport": {"width": settings.viewport_width, "height": settings.viewport_height},
        "device_scale_factor": settings.device_scale_factor,
    }

    if settings.auth_mode == "integrated":
        auth_allowlist = build_auth_server_allowlist(settings)
        launch_args = []
        if auth_allowlist:
            launch_args.extend(
                [
                    f"--auth-server-whitelist={auth_allowlist}",
                    f"--auth-server-allowlist={auth_allowlist}",
                    f"--auth-negotiate-delegate-whitelist={auth_allowlist}",
                    f"--auth-negotiate-delegate-allowlist={auth_allowlist}",
                    "--auth-schemes=basic,digest,ntlm,negotiate",
                ]
            )
        if settings.edge_profile_directory:
            launch_args.append(f"--profile-directory={settings.edge_profile_directory}")

        user_data_dir = (
            Path(settings.edge_user_data_dir).expanduser()
            if settings.edge_user_data_dir
            else settings.browser_profile_dir.resolve()
        )
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(user_data_dir),
            headless=settings.headless,
            channel=settings.browser_channel,
            args=launch_args,
            **common_context_args,
        )
        page = context.pages[0] if context.pages else context.new_page()
        return context, page

    browser = playwright.chromium.launch(
        headless=settings.headless,
        channel=settings.browser_channel,
    )

    if settings.auth_mode == "basic":
        if not settings.http_username or not settings.http_password:
            raise ValueError("HTTP_USERNAME and HTTP_PASSWORD are required for AUTH_MODE=basic.")
        common_context_args["http_credentials"] = {
            "username": settings.http_username,
            "password": settings.http_password,
        }

    context = browser.new_context(**common_context_args)
    page = context.new_page()
    return context, page


def first_visible_selector(page: Page, selectors: list[str]) -> str | None:
    for selector in selectors:
        try:
            locator = page.locator(selector).first
            locator.wait_for(state="visible", timeout=1000)
            return selector
        except Exception:
            continue
    return None


def handle_login_form_if_needed(page: Page, settings: Settings, logger: logging.Logger) -> None:
    username_selector = settings.login_username_selector or first_visible_selector(page, USERNAME_SELECTORS)
    password_selector = settings.login_password_selector or first_visible_selector(page, PASSWORD_SELECTORS)
    if not username_selector or not password_selector:
        return

    if not settings.pbirs_username or not settings.pbirs_password:
        raise ValueError(
            "A login form was detected, but PBIRS_USERNAME/PBIRS_PASSWORD are not configured."
        )

    logger.info("Login form detected. Submitting configured credentials.")
    page.locator(username_selector).first.fill(settings.pbirs_username)
    page.locator(password_selector).first.fill(settings.pbirs_password)

    submit_selector = settings.login_submit_selector or first_visible_selector(page, SUBMIT_SELECTORS)
    if submit_selector:
        page.locator(submit_selector).first.click()
    else:
        page.locator(password_selector).first.press("Enter")

    page.wait_for_load_state("networkidle", timeout=settings.navigation_timeout_ms)


def frame_surface_snapshot(frame: Any) -> dict[str, Any]:
    return frame.evaluate(
        """
        ({ visualSelectors, spinnerSelectors, tabSelectors }) => {
          const visible = (el) => {
            if (!el) return false;
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style &&
              style.display !== 'none' &&
              style.visibility !== 'hidden' &&
              Number(style.opacity || 1) > 0.01 &&
              rect.width > 4 &&
              rect.height > 4;
          };

          const countVisible = (selectors) => selectors.reduce((total, selector) => {
            return total + Array.from(document.querySelectorAll(selector)).filter(visible).length;
          }, 0);

          const text = (document.body?.innerText || '').toLowerCase();
          return {
            href: window.location.href,
            title: document.title,
            hasPowerBiApi: Boolean(window.powerbi),
            visualCount: countVisible(visualSelectors),
            spinnerCount: countVisible(spinnerSelectors),
            tabCount: countVisible(tabSelectors),
            unauthorizedText: text.includes('unauthorized'),
            signInText: text.includes('sign in') || text.includes('login') || text.includes('log in')
          };
        }
        """,
        {
            "visualSelectors": VISUAL_SELECTORS,
            "spinnerSelectors": SPINNER_SELECTORS,
            "tabSelectors": TAB_SELECTORS,
        },
    )


def locate_report_frame(page: Page, settings: Settings, logger: logging.Logger) -> Any:
    deadline = time.monotonic() + (settings.report_render_timeout_ms / 1000)
    best_frame = page.main_frame
    best_score = -1

    while time.monotonic() < deadline:
        for frame in page.frames:
            try:
                snapshot = frame_surface_snapshot(frame)
            except Exception:
                continue

            score = (
                snapshot["visualCount"] * 10
                + snapshot["tabCount"] * 5
                + (50 if snapshot["hasPowerBiApi"] else 0)
                - snapshot["spinnerCount"]
            )
            if score > best_score:
                best_score = score
                best_frame = frame

            if snapshot["hasPowerBiApi"] or snapshot["visualCount"] > 0:
                logger.info(
                    "Selected report frame: href=%s | visuals=%s | tabs=%s | powerbi_api=%s",
                    snapshot["href"],
                    snapshot["visualCount"],
                    snapshot["tabCount"],
                    snapshot["hasPowerBiApi"],
                )
                return frame
        time.sleep(1)

    logger.info("Falling back to the highest-scoring frame with score=%s", best_score)
    return best_frame


def wait_for_report_ready(frame: Any, settings: Settings, logger: logging.Logger, label: str) -> None:
    deadline = time.monotonic() + (settings.report_render_timeout_ms / 1000)
    stable_hits = 0
    previous_signature: tuple[int, int, int, bool] | None = None

    while time.monotonic() < deadline:
        snapshot = frame_surface_snapshot(frame)
        signature = (
            snapshot["visualCount"],
            snapshot["spinnerCount"],
            snapshot["tabCount"],
            snapshot["hasPowerBiApi"],
        )

        renderable = (
            snapshot["spinnerCount"] == 0
            and (snapshot["visualCount"] > 0 or snapshot["hasPowerBiApi"] or snapshot["tabCount"] > 0)
        )

        if renderable and signature == previous_signature:
            stable_hits += 1
        elif renderable:
            stable_hits = 1
        else:
            stable_hits = 0

        if stable_hits >= settings.report_stable_polls:
            logger.info(
                "Report surface is stable for '%s' | visuals=%s | tabs=%s",
                label,
                snapshot["visualCount"],
                snapshot["tabCount"],
            )
            return

        previous_signature = signature
        time.sleep(settings.report_stable_interval_ms / 1000)

    raise TimeoutError(f"Timed out waiting for report rendering to stabilize for '{label}'.")


def discover_tabs_via_api(frame: Any, logger: logging.Logger) -> list[ReportTab]:
    raw_pages = frame.evaluate(
        """
        async () => {
          const getReport = () => {
            if (!window.powerbi) {
              return null;
            }

            if (Array.isArray(window.powerbi.embeds)) {
              for (const embed of window.powerbi.embeds) {
                if (embed && typeof embed.getPages === 'function' && typeof embed.setPage === 'function') {
                  return embed;
                }
              }
            }

            if (typeof window.powerbi.get === 'function') {
              for (const el of document.querySelectorAll('*')) {
                try {
                  const embed = window.powerbi.get(el);
                  if (embed && typeof embed.getPages === 'function' && typeof embed.setPage === 'function') {
                    return embed;
                  }
                } catch (error) {
                  // Ignore nodes that are not Power BI containers.
                }
              }
            }

            return null;
          };

          const report = getReport();
          if (!report) {
            return [];
          }

          const pages = await report.getPages();
          let activePageName = null;

          try {
            const activePage = await report.getActivePage();
            activePageName = activePage?.name || null;
          } catch (error) {
            activePageName = null;
          }

          return pages.map((page, index) => ({
            index,
            pageName: page.name || `page-${index + 1}`,
            label: page.displayName || page.name || `Page ${index + 1}`,
            isActive: activePageName ? activePageName === page.name : Boolean(page.isActive),
          }));
        }
        """
    )

    tabs = [
        ReportTab(
            label=str(item["label"]).strip() or f"Page {item['index'] + 1}",
            mode="api",
            page_name=item["pageName"],
            is_active=bool(item["isActive"]),
        )
        for item in raw_pages
    ]
    logger.info("Discovered %s tabs through the Power BI API: %s", len(tabs), [tab.label for tab in tabs])
    return tabs


def discover_tabs_via_dom(frame: Any, settings: Settings, logger: logging.Logger) -> list[ReportTab]:
    thumbnail_tabs = frame.evaluate(
        """
        (selectors) => {
          const visible = (el) => {
            if (!el) return false;
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style &&
              style.display !== 'none' &&
              style.visibility !== 'hidden' &&
              Number(style.opacity || 1) > 0.01 &&
              rect.width > 20 &&
              rect.height > 20;
          };

          const seen = new Set();
          const tabs = [];

          for (const selector of selectors) {
            for (const el of document.querySelectorAll(selector)) {
              if (!visible(el)) continue;
              if (el.classList.contains('hidden-tab')) continue;

              const label = (el.innerText || el.textContent || '').replace(/\\s+/g, ' ').trim();
              if (!label) continue;

              const rect = el.getBoundingClientRect();
              const key = `${label}|${Math.round(rect.x)}|${Math.round(rect.y)}`;
              if (seen.has(key)) continue;
              seen.add(key);

              tabs.push({
                label,
                isActive: el.classList.contains('selected') || el.getAttribute('aria-selected') === 'true',
                x: rect.x,
                y: rect.y,
                domIndex: tabs.length
              });
            }
          }

          tabs.sort((a, b) => (a.y - b.y) || (a.x - b.x));
          return tabs;
        }
        """,
        PAGE_THUMBNAIL_SELECTORS,
    )

    if thumbnail_tabs:
        tabs = [
            ReportTab(
                label=item["label"],
                mode="dom",
                dom_index=int(item["domIndex"]),
                is_active=bool(item["isActive"]),
            )
            for item in thumbnail_tabs
        ]
        if settings.expected_sheets:
            expected_map = {normalize_sheet_name(name): name for name in settings.expected_sheets}
            filtered = [tab for tab in tabs if normalize_sheet_name(tab.label) in expected_map]
            if filtered:
                order = {normalize_sheet_name(name): index for index, name in enumerate(settings.expected_sheets)}
                filtered.sort(key=lambda tab: order.get(normalize_sheet_name(tab.label), 999))
                tabs = filtered
        logger.info("Discovered %s sheet thumbnails through DOM inspection: %s", len(tabs), [tab.label for tab in tabs])
        return tabs

    raw_tabs = frame.evaluate(
        """
        (tabSelectors) => {
          const visible = (el) => {
            if (!el) return false;
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style &&
              style.display !== 'none' &&
              style.visibility !== 'hidden' &&
              Number(style.opacity || 1) > 0.01 &&
              rect.width > 4 &&
              rect.height > 4;
          };

          const seen = new Set();
          const groups = new Map();

          const classify = (el) => {
            const classText = `${el.className || ''} ${(el.parentElement && el.parentElement.className) || ''}`.toLowerCase();
            const attrText = `${el.getAttribute('aria-label') || ''} ${el.getAttribute('title') || ''}`.toLowerCase();
            const combined = `${classText} ${attrText}`;

            if (combined.includes('page') || combined.includes('sheet') || combined.includes('tab')) {
              return 3;
            }
            if (combined.includes('pivot') || combined.includes('nav')) {
              return 2;
            }
            return 0;
          };

          const pushTab = (el) => {
            if (!visible(el)) return;

            const label =
              el.getAttribute('aria-label') ||
              el.getAttribute('title') ||
              el.textContent ||
              '';
            const normalized = label.replace(/\\s+/g, ' ').trim();
            if (!normalized) return;

            const rect = el.getBoundingClientRect();
            const key = `${normalized}|${Math.round(rect.x)}|${Math.round(rect.y)}`;
            if (seen.has(key)) return;
            seen.add(key);

            const isActive =
              el.getAttribute('aria-selected') === 'true' ||
              el.getAttribute('tabindex') === '0' ||
              el.classList.contains('active') ||
              el.classList.contains('is-active');

            const yBucket = Math.round(rect.y / 24) * 24;
            const bucketKey = `${yBucket}`;
            if (!groups.has(bucketKey)) {
              groups.set(bucketKey, []);
            }

            groups.get(bucketKey).push({
              label: normalized,
              isActive,
              x: rect.x,
              y: rect.y,
              width: rect.width,
              height: rect.height,
              score: classify(el)
            });
          };

          for (const selector of tabSelectors) {
            for (const el of document.querySelectorAll(selector)) {
              pushTab(el);
            }
          }

          const groupedTabs = Array.from(groups.values())
            .map((tabs) => tabs.sort((a, b) => a.x - b.x))
            .filter((tabs) => tabs.length >= 2)
            .map((tabs) => {
              const uniqueLabels = new Set(tabs.map((tab) => tab.label));
              const activeCount = tabs.filter((tab) => tab.isActive).length;
              const score =
                tabs.reduce((sum, tab) => sum + tab.score, 0) * 100 +
                uniqueLabels.size * 10 -
                Math.abs(uniqueLabels.size - 6) * 5 -
                Math.abs(activeCount - 1) * 20;
              return { tabs, score, y: tabs[0].y };
            })
            .sort((a, b) => b.score - a.score || a.y - b.y);

          const chosen = groupedTabs.length > 0 ? groupedTabs[0].tabs : [];
          return chosen.map((tab, index) => ({
            ...tab,
            domIndex: index
          }));
        }
        """,
        TAB_SELECTORS,
    )

    tabs = [
        ReportTab(
            label=item["label"],
            mode="dom",
            dom_index=int(item["domIndex"]),
            is_active=bool(item["isActive"]),
        )
        for item in raw_tabs
    ]
    if settings.expected_sheets:
        expected_map = {normalize_sheet_name(name): name for name in settings.expected_sheets}
        filtered = [tab for tab in tabs if normalize_sheet_name(tab.label) in expected_map]
        if filtered:
            order = {normalize_sheet_name(name): index for index, name in enumerate(settings.expected_sheets)}
            filtered.sort(key=lambda tab: order.get(normalize_sheet_name(tab.label), 999))
            tabs = filtered

    logger.info("Discovered %s tabs through DOM inspection: %s", len(tabs), [tab.label for tab in tabs])
    return tabs


def discover_tabs(frame: Any, settings: Settings, logger: logging.Logger) -> tuple[list[ReportTab], Any | None]:
    try:
        tabs = discover_tabs_via_api(frame, logger)
        if tabs:
            if settings.expected_sheets:
                order = {normalize_sheet_name(name): index for index, name in enumerate(settings.expected_sheets)}
                filtered = [tab for tab in tabs if normalize_sheet_name(tab.label) in order]
                if filtered:
                    filtered.sort(key=lambda tab: order.get(normalize_sheet_name(tab.label), 999))
                    tabs = filtered
            return tabs, "api"
    except Exception as error:
        logger.warning("Power BI API tab discovery failed: %s", error)

    tabs = discover_tabs_via_dom(frame, settings, logger)
    return tabs, None


def activate_tab(frame: Any, report_handle: Any | None, tab: ReportTab, settings: Settings, logger: logging.Logger) -> None:
    logger.info("Opening tab '%s' using %s mode.", tab.label, tab.mode)

    if tab.mode == "api":
        if not tab.page_name:
            raise RuntimeError(f"Missing Power BI page name for tab '{tab.label}'.")
        frame.evaluate(
            """
            async (pageName) => {
              const getReport = () => {
                if (!window.powerbi) {
                  return null;
                }

                if (Array.isArray(window.powerbi.embeds)) {
                  for (const embed of window.powerbi.embeds) {
                    if (embed && typeof embed.setPage === 'function') {
                      return embed;
                    }
                  }
                }

                if (typeof window.powerbi.get === 'function') {
                  for (const el of document.querySelectorAll('*')) {
                    try {
                      const embed = window.powerbi.get(el);
                      if (embed && typeof embed.setPage === 'function') {
                        return embed;
                      }
                    } catch (error) {
                      // Ignore non-report nodes.
                    }
                  }
                }

                return null;
              };

              const report = getReport();
              if (!report) {
                throw new Error('Power BI report object was not found while switching pages.');
              }
              await report.setPage(pageName);
            }
            """,
            tab.page_name,
        )
    else:
        clicked = frame.evaluate(
            """
            ({ pageThumbnailSelectors, tabSelectors, targetLabel, targetIndex }) => {
              const visible = (el) => {
                if (!el) return false;
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style &&
                  style.display !== 'none' &&
                  style.visibility !== 'hidden' &&
                  Number(style.opacity || 1) > 0.01 &&
                  rect.width > 4 &&
                  rect.height > 4;
              };

              const normalizedTarget = (targetLabel || '').replace(/\\s+/g, ' ').trim();

              const clickElement = (el) => {
                el.scrollIntoView({ block: 'center', inline: 'center' });
                if (typeof el.click === 'function') {
                  el.click();
                  return true;
                }
                el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
                return true;
              };

              const thumbnailCandidates = [];
              const seenThumbs = new Set();
              for (const selector of pageThumbnailSelectors) {
                for (const el of document.querySelectorAll(selector)) {
                  if (!visible(el)) continue;
                  if (el.classList.contains('hidden-tab')) continue;
                  const label = (el.innerText || el.textContent || '').replace(/\\s+/g, ' ').trim();
                  if (!label) continue;
                  const rect = el.getBoundingClientRect();
                  const key = `${label}|${Math.round(rect.x)}|${Math.round(rect.y)}`;
                  if (seenThumbs.has(key)) continue;
                  seenThumbs.add(key);
                  thumbnailCandidates.push({ el, label, x: rect.x, y: rect.y });
                }
              }

              thumbnailCandidates.sort((a, b) => (a.y - b.y) || (a.x - b.x));
              let thumbnailMatch = thumbnailCandidates.find((item) => item.label === normalizedTarget) || null;
              if (!thumbnailMatch && Number.isInteger(targetIndex) && targetIndex >= 0 && targetIndex < thumbnailCandidates.length) {
                thumbnailMatch = thumbnailCandidates[targetIndex];
              }
              if (thumbnailMatch) {
                return clickElement(thumbnailMatch.el);
              }

              const tabs = [];
              const seen = new Set();
              for (const selector of tabSelectors) {
                for (const el of document.querySelectorAll(selector)) {
                  if (!visible(el)) continue;
                  const label =
                    el.getAttribute('aria-label') ||
                    el.getAttribute('title') ||
                    el.textContent ||
                    '';
                  const normalized = label.replace(/\\s+/g, ' ').trim();
                  if (!normalized) continue;
                  const rect = el.getBoundingClientRect();
                  const key = `${normalized}|${Math.round(rect.x)}|${Math.round(rect.y)}`;
                  if (seen.has(key)) continue;
                  seen.add(key);
                  tabs.push({ el, label: normalized, x: rect.x, y: rect.y });
                }
              }

              tabs.sort((a, b) => (a.y - b.y) || (a.x - b.x));
              let candidate = tabs.find((item) => item.label === normalizedTarget) || null;
              if (!candidate && Number.isInteger(targetIndex) && targetIndex >= 0 && targetIndex < tabs.length) {
                candidate = tabs[targetIndex];
              }
              if (!candidate) {
                return false;
              }
              return clickElement(candidate.el);
            }
            """,
            {
                "pageThumbnailSelectors": PAGE_THUMBNAIL_SELECTORS,
                "tabSelectors": TAB_SELECTORS,
                "targetLabel": tab.label,
                "targetIndex": tab.dom_index,
            },
        )
        if not clicked:
            raise RuntimeError(f"Could not locate DOM tab '{tab.label}' for activation.")

    time.sleep(settings.post_tab_click_wait_ms / 1000)
    wait_for_report_ready(frame, settings, logger, tab.label)


def get_slicer_options(frame: Any, settings: Settings, logger: logging.Logger) -> list[str]:
    """Open the slicer dropdown, list all options, close it, and filter by exclusion list."""
    slicer_name = settings.filter_slicer_name
    if not slicer_name:
        return []
    
    logger.info("Opening slicer '%s' dropdown to extract options...", slicer_name)
    try:
        # Locate dropdown menu inside the slicer
        slicer_dropdown = frame.locator(f".slicer-container:has-text('{slicer_name}') .slicer-dropdown-menu")
        slicer_dropdown.click()
        
        # Wait for options to render
        frame.locator(".slicerText").first.wait_for(state="visible", timeout=10000)
        
        # Extract option texts
        options_elements = frame.locator(".slicerText").all()
        options = [el.inner_text().strip() for el in options_elements if el.inner_text().strip()]
        
        # Close the dropdown
        slicer_dropdown.click()
        
        # Filter options
        exclude = set(settings.filter_exclude_options)
        filtered_options = [opt for opt in options if opt not in exclude]
        
        logger.info("Retrieved %d slicer options (filtered: %s)", len(filtered_options), filtered_options)
        return filtered_options
    except Exception as error:
        logger.exception("Failed to extract slicer options for '%s': %s", slicer_name, error)
        raise


def select_slicer_option(frame: Any, option_text: str, settings: Settings, logger: logging.Logger) -> None:
    """Clear selections, open the slicer dropdown, click the target option, and close it.
    Uses a forced click to bypass overlay interception issues.
    """
    slicer_name = settings.filter_slicer_name
    if not slicer_name:
        return

    logger.info("Selecting slicer option '%s' on slicer '%s'", option_text, slicer_name)
    try:
        # 1. Clear previous selections if clear button is visible
        clear_btn = frame.locator(f".slicer-container:has-text('{slicer_name}') .slicer-header-clear")
        if clear_btn.is_visible():
            clear_btn.click(force=True)
            logger.info("Cleared existing selections for slicer '%s'", slicer_name)

        # 2. Open dropdown
        slicer_dropdown = frame.locator(f".slicer-container:has-text('{slicer_name}') .slicer-dropdown-menu")
        slicer_dropdown.click(force=True)

        # 3. Locate the target option
        option_locator = frame.locator(f".slicerText:has-text('{option_text}')").first
        option_locator.wait_for(state="visible", timeout=10000)
        # Ensure the element is in view before clicking
        option_locator.scroll_into_view_if_needed()
        # Use forced click to avoid overlay interception
        option_locator.click(force=True)
        logger.info("Clicked option '%s'", option_text)

        # 4. Close dropdown
        slicer_dropdown.click(force=True)

        # Allow the report to stabilize after changing the filter
        time.sleep(settings.post_tab_click_wait_ms / 1000)
        wait_for_report_ready(frame, settings, logger, f"option select: {option_text}")
    except Exception as error:
        logger.exception(
            "Failed to select slicer option '%s' for slicer '%s': %s",
            option_text,
            slicer_name,
            error,
        )
        raise


def capture_tab_screenshot(page: Page, frame: Any, output_path: Path) -> None:
    dimensions = frame.evaluate(
        """
        () => ({
          width: Math.max(
            document.documentElement?.scrollWidth || 0,
            document.body?.scrollWidth || 0,
            document.documentElement?.clientWidth || 0,
            document.body?.clientWidth || 0
          ),
          height: Math.max(
            document.documentElement?.scrollHeight || 0,
            document.body?.scrollHeight || 0,
            document.documentElement?.clientHeight || 0,
            document.body?.clientHeight || 0
          )
        })
        """
    )

    frame_element = frame.frame_element()
    frame_element.evaluate(
        """
        (el, dims) => {
          el.style.width = `${Math.ceil(dims.width)}px`;
          el.style.height = `${Math.ceil(dims.height)}px`;
          el.style.maxWidth = 'none';
          el.style.maxHeight = 'none';
          el.style.border = '0';
          el.style.display = 'block';
        }
        """,
        dimensions,
    )

    frame.evaluate("() => window.scrollTo(0, 0)")
    page.evaluate("() => window.scrollTo(0, 0)")
    page.screenshot(
        path=str(output_path),
        full_page=True,
        animations="disabled",
    )


def capture_report(settings: Settings, logger: logging.Logger) -> tuple[list[Path], list[str]]:
    screenshots: list[Path] = []
    errors: list[str] = []

    if sync_playwright is None:
        raise RuntimeError(
            "Playwright is not installed in this environment. "
            "Run '.\\.venv\\Scripts\\python.exe -m pip install -r requirements.txt' "
            "and then '.\\.venv\\Scripts\\python.exe -m playwright install chromium'."
        )

    with sync_playwright() as playwright:
        context, page = build_context(playwright, settings)
        try:
            page.set_default_timeout(settings.navigation_timeout_ms)
            page.set_default_navigation_timeout(settings.navigation_timeout_ms)

            logger.info("Opening report URL: %s", settings.report_url)
            try:
                response = page.goto(settings.report_url, wait_until="domcontentloaded")
            except PlaywrightError as error:
                if "ERR_INVALID_AUTH_CREDENTIALS" in str(error):
                    raise RuntimeError(
                        "PBIRS authentication failed before the page loaded. "
                        "If AUTH_MODE=integrated, run the script under a Windows account that already has access "
                        "to the report and keep PBIRS_USERNAME/PBIRS_PASSWORD empty. "
                        "If the server uses a login form instead, set AUTH_MODE=form."
                    ) from error
                raise
            if response is not None:
                logger.info("Initial response status: %s", response.status)
                if response.status == 401 and settings.auth_mode == "none":
                    raise PermissionError(
                        "The report returned HTTP 401. Configure AUTH_MODE and credentials before running again."
                    )

            page.wait_for_load_state("networkidle", timeout=settings.navigation_timeout_ms)
            handle_login_form_if_needed(page, settings, logger)

            report_frame = locate_report_frame(page, settings, logger)
            wait_for_report_ready(report_frame, settings, logger, "initial load")
            tabs, report_handle = discover_tabs(report_frame, settings, logger)

            if not tabs:
                logger.warning("No tabs were discovered. Capturing the current report surface as a single page.")
                tabs = [ReportTab(label="current_view", mode="dom", dom_index=0, is_active=True)]

            if settings.filter_slicer_name:
                # Filter by option flow
                slicer_tab = None
                if settings.filter_slicer_page:
                    slicer_tab = next(
                        (t for t in tabs if t.label.strip().casefold() == settings.filter_slicer_page.strip().casefold()),
                        None
                    )
                if not slicer_tab:
                    logger.warning("Slicer page '%s' not found in discovered tabs. Using first tab.", settings.filter_slicer_page)
                    slicer_tab = tabs[0]
                
                # Activate the slicer page to extract options
                logger.info("Activating slicer page '%s' to retrieve options...", slicer_tab.label)
                activate_tab(report_frame, report_handle, slicer_tab, settings, logger)
                
                options = get_slicer_options(report_frame, settings, logger)
                if not options:
                    logger.warning("No filter options found. Capturing standard report tabs without filtering.")
                    options = [None]
                
                run_stamp = timestamp_compact(settings.timezone)
                dated_output_dir = settings.output_dir / run_stamp[:8]
                dated_output_dir.mkdir(parents=True, exist_ok=True)
                
                screenshot_idx = 1
                for option in options:
                    if option:
                        # Return to the slicer page to switch filter
                        logger.info("Switching to slicer page '%s' to select option '%s'...", slicer_tab.label, option)
                        activate_tab(report_frame, report_handle, slicer_tab, settings, logger)
                        select_slicer_option(report_frame, option, settings, logger)
                        
                    for tab in tabs:
                        try:
                            # Activate the target tab
                            activate_tab(report_frame, report_handle, tab, settings, logger)
                            
                            # Construct filename
                            opt_segment = f"{sanitize_filename(option)}_" if option else ""
                            filename = f"{screenshot_idx:02d}_{opt_segment}{sanitize_filename(tab.label)}_{run_stamp}.png"
                            output_path = dated_output_dir / filename
                            
                            capture_tab_screenshot(page, report_frame, output_path)
                            screenshots.append(output_path)
                            logger.info("Saved screenshot: %s", output_path)
                            screenshot_idx += 1
                        except Exception as error:
                            opt_msg = f" (Option: '{option}')" if option else ""
                            message = f"Tab '{tab.label}' failed{opt_msg}: {error}"
                            logger.exception(message)
                            errors.append(message)
            else:
                # Original tab-by-tab flow
                run_stamp = timestamp_compact(settings.timezone)
                dated_output_dir = settings.output_dir / run_stamp[:8]
                dated_output_dir.mkdir(parents=True, exist_ok=True)

                for index, tab in enumerate(tabs, start=1):
                    try:
                        if not (index == 1 and tab.is_active):
                            activate_tab(report_frame, report_handle, tab, settings, logger)
                        else:
                            wait_for_report_ready(report_frame, settings, logger, tab.label)

                        filename = f"{index:02d}_{sanitize_filename(tab.label)}_{run_stamp}.png"
                        output_path = dated_output_dir / filename
                        capture_tab_screenshot(page, report_frame, output_path)
                        screenshots.append(output_path)
                        logger.info("Saved screenshot: %s", output_path)
                    except Exception as error:
                        message = f"Tab '{tab.label}' failed: {error}"
                        logger.exception(message)
                        errors.append(message)
        finally:
            context.close()

    return screenshots, errors


def _extract_sheet_label(filename: str) -> str:
    """Extract a human-readable sheet label from a screenshot filename.

    Filenames look like ``01_Home_20260503_110000.png``.
    This returns ``Home`` (dropping the numeric prefix and the timestamp suffix).
    """
    stem = Path(filename).stem                       # 01_Home_20260503_110000
    parts = stem.split("_")
    # Drop leading numeric index and trailing timestamp segments (YYYYMMDD, HHMMSS)
    meaningful = []
    for part in parts:
        if not part:
            continue
        # Skip pure-digit segments (index and timestamp chunks)
        if part.isdigit():
            continue
        meaningful.append(part)
    return " ".join(meaningful) if meaningful else stem


def _build_html_email(
    settings: Settings,
    stamp: str,
    attachments: list[Path],
    errors: list[str],
) -> str:
    """Build a premium HTML email body with inline CID references for each screenshot."""

    # ── per-figure HTML blocks ──
    figure_blocks: list[str] = []
    for index, attachment in enumerate(attachments, start=1):
        label = _extract_sheet_label(attachment.name)
        cid = f"screenshot_{index}"
        figure_blocks.append(
            f"""
            <!-- Figure {index} -->
            <tr>
              <td style="padding: 0 32px 28px 32px;">
                <table width="100%" cellpadding="0" cellspacing="0" border="0"
                       style="background-color: #ffffff; border-radius: 12px;
                              border: 1px solid #e2e8f0; overflow: hidden;">
                  <!-- Figure header -->
                  <tr>
                    <td style="padding: 14px 20px; background: linear-gradient(135deg, #f0f4ff 0%, #e8f0fe 100%);
                               border-bottom: 1px solid #e2e8f0;">
                      <table cellpadding="0" cellspacing="0" border="0">
                        <tr>
                          <td style="background-color: #4a6cf7; color: #ffffff; font-size: 12px;
                                     font-weight: 700; padding: 4px 10px; border-radius: 6px;
                                     font-family: 'Segoe UI', Arial, sans-serif; letter-spacing: 0.3px;">
                            {index}
                          </td>
                          <td style="padding-left: 12px; font-size: 15px; font-weight: 600;
                                     color: #1e293b; font-family: 'Segoe UI', Arial, sans-serif;">
                            {label}
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                  <!-- Screenshot image -->
                  <tr>
                    <td style="padding: 16px;">
                      <img src="cid:{cid}" alt="Illustration {index} – {label}"
                           width="900"
                            style="display:block; width:100%; 
                            max-width:900px; height:auto; margin:0 auto;"/>
                    </td>
                  </tr>
                  <!-- Figure caption -->
                  <tr>
                    <td style="padding: 0 20px 14px 20px; font-size: 12px; color: #94a3b8;
                               font-family: 'Segoe UI', Arial, sans-serif; text-align: center;">
                      Figure {index} &mdash; {label}
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
            """
        )

    figures_html = "\n".join(figure_blocks)

    # ── error section (if any) ──
    errors_html = ""
    if errors:
        error_items = "".join(
            f'<li style="padding: 4px 0; color: #dc2626; font-size: 13px;">{err}</li>'
            for err in errors
        )
        errors_html = f"""
        <tr>
          <td style="padding: 0 32px 28px 32px;">
            <table width="100%" cellpadding="0" cellspacing="0" border="0"
                   style="background-color: #fef2f2; border-radius: 12px;
                          border: 1px solid #fecaca;">
              <tr>
                <td style="padding: 16px 20px;">
                  <p style="margin: 0 0 8px 0; font-size: 14px; font-weight: 600;
                            color: #991b1b; font-family: 'Segoe UI', Arial, sans-serif;">
                    ⚠ Erreurs de capture
                  </p>
                  <ul style="margin: 0; padding-left: 18px;
                             font-family: 'Segoe UI', Arial, sans-serif;">
                    {error_items}
                  </ul>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        """

    status_badge = (
        '<span style="background-color: #fbbf24; color: #78350f; font-size: 11px; '
        'font-weight: 700; padding: 3px 10px; border-radius: 20px; '
        'letter-spacing: 0.5px;">PARTIEL</span>'
        if errors
        else '<span style="background-color: #34d399; color: #064e3b; font-size: 11px; '
        'font-weight: 700; padding: 3px 10px; border-radius: 20px; '
        'letter-spacing: 0.5px;">SUCCÈS</span>'
    )

    return f"""\
<!DOCTYPE html>
<html lang="fr" xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>{settings.email_subject_prefix}</title>
</head>
<body style="margin: 0; padding: 0; background-color: #f1f5f9;
             font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
             -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%;">
  <!-- Outer wrapper -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0"
         style="background-color: #f1f5f9; padding: 24px 0;">
    <tr>
      <td align="center">
        <!-- Main card -->
        <table width="680" cellpadding="0" cellspacing="0" border="0"
               style="max-width: 680px; width: 100%; background-color: #ffffff;
                      border-radius: 16px; overflow: hidden;
                      box-shadow: 0 4px 24px rgba(0,0,0,0.06);">

          <!-- ═══ Header ═══ -->
          <tr>
            <td style="background: linear-gradient(135deg, #4a6cf7 0%, #6366f1 50%, #8b5cf6 100%);
                       padding: 36px 32px 28px 32px; text-align: center;">
              <p style="margin: 0 0 6px 0; font-size: 28px; font-weight: 700;
                        color: #ffffff; letter-spacing: -0.5px;
                        font-family: 'Segoe UI', Arial, sans-serif;">
                Capture Quotidienne des Rapports
              </p>
              <p style="margin: 0; font-size: 14px; color: rgba(255,255,255,0.85);
                        font-family: 'Segoe UI', Arial, sans-serif;">
                {settings.email_subject_prefix}
              </p>
            </td>
          </tr>

          <!-- ═══ Summary bar ═══ -->
          <tr>
            <td style="padding: 24px 32px 20px 32px;">
              <table width="100%" cellpadding="0" cellspacing="0" border="0"
                     style="background-color: #f8fafc; border-radius: 12px;
                            border: 1px solid #e2e8f0;">
                <tr>
                  <!-- Status -->
                  <td style="padding: 16px 20px; text-align: center; width: 33%;
                             border-right: 1px solid #e2e8f0;">
                    <p style="margin: 0 0 4px 0; font-size: 11px; font-weight: 600;
                              color: #94a3b8; text-transform: uppercase;
                              letter-spacing: 0.8px;
                              font-family: 'Segoe UI', Arial, sans-serif;">
                      Statut
                    </p>
                    {status_badge}
                  </td>
                  <!-- Sheets -->
                  <td style="padding: 16px 20px; text-align: center; width: 33%;
                             border-right: 1px solid #e2e8f0;">
                    <p style="margin: 0 0 4px 0; font-size: 11px; font-weight: 600;
                              color: #94a3b8; text-transform: uppercase;
                              letter-spacing: 0.8px;
                              font-family: 'Segoe UI', Arial, sans-serif;">
                      Feuilles
                    </p>
                    <p style="margin: 0; font-size: 22px; font-weight: 700;
                              color: #1e293b;
                              font-family: 'Segoe UI', Arial, sans-serif;">
                      {len(attachments)}
                    </p>
                  </td>
                  <!-- Timestamp -->
                  <td style="padding: 16px 20px; text-align: center; width: 34%;">
                    <p style="margin: 0 0 4px 0; font-size: 11px; font-weight: 600;
                              color: #94a3b8; text-transform: uppercase;
                              letter-spacing: 0.8px;
                              font-family: 'Segoe UI', Arial, sans-serif;">
                      Capturé le
                    </p>
                    <p style="margin: 0; font-size: 13px; font-weight: 600;
                              color: #475569;
                              font-family: 'Segoe UI', Arial, sans-serif;">
                      {stamp}
                    </p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- ═══ Section title ═══ -->
          <tr>
            <td style="padding: 8px 32px 16px 32px;">
              <p style="margin: 0; font-size: 18px; font-weight: 700; color: #1e293b;
                        font-family: 'Segoe UI', Arial, sans-serif;">
                Captures d'écran du rapport
              </p>
              <p style="margin: 4px 0 0 0; font-size: 13px; color: #64748b;
                        font-family: 'Segoe UI', Arial, sans-serif;">
                Chaque feuille capturée est affichée ci-dessous avec son numéro de figure correspondant.
              </p>
            </td>
          </tr>

          <!-- ═══ Figures ═══ -->
          {figures_html}

          <!-- ═══ Errors (if any) ═══ -->
          {errors_html}

          <!-- ═══ Footer ═══ -->
          <tr>
            <td style="padding: 24px 32px; background-color: #f8fafc;
                       border-top: 1px solid #e2e8f0; text-align: center;">
              <p style="margin: 0 0 4px 0; font-size: 12px; color: #94a3b8;
                        font-family: 'Segoe UI', Arial, sans-serif;">
                Rapport automatisé généré par <strong style="color: #64748b;">BI Bot</strong>
              </p>
              <p style="margin: 0; font-size: 11px; color: #cbd5e1;
                        font-family: 'Segoe UI', Arial, sans-serif;">
                Ceci est un e-mail automatique. Merci de ne pas répondre directement.
              </p>
            </td>
          </tr>

        </table>
        <!-- /Main card -->
      </td>
    </tr>
  </table>
</body>
</html>"""


def send_email(settings: Settings, logger: logging.Logger, attachments: list[Path], errors: list[str]) -> None:
    require_email_config(settings)
    if not attachments:
        raise ValueError("No screenshots were generated, so the email was not sent.")

    stamp = timestamp_readable(settings.timezone)
    status_prefix = "[PARTIAL] " if errors else ""
    subject = f"{status_prefix}{settings.email_subject_prefix} - {stamp}"

    html_body = _build_html_email(settings, stamp, attachments, errors)

    # ── plain-text fallback ──
    plain_lines = [
        f"{settings.email_subject_prefix}",
        f"Horodatage : {stamp}",
        f"URL du rapport : {settings.report_url}",
        f"Feuilles capturées : {len(attachments)}",
        "",
    ]
    for index, attachment in enumerate(attachments, start=1):
        label = _extract_sheet_label(attachment.name)
        plain_lines.append(f"  {index}. {label}  (voir pièce jointe : {attachment.name})")
    if errors:
        plain_lines.append("")
        plain_lines.append("Erreurs :")
        plain_lines.extend(f"  - {error}" for error in errors)

    # ── build MIME message ──
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.image import MIMEImage

    msg = MIMEMultipart("related")
    msg["Subject"] = subject
    msg["From"] = settings.email_from
    msg["To"] = ", ".join(settings.email_to)
    if settings.email_reply_to:
        msg["Reply-To"] = settings.email_reply_to

    # Attach the HTML + plain-text alternative
    alt_part = MIMEMultipart("alternative")
    alt_part.attach(MIMEText("\n".join(plain_lines), "plain", "utf-8"))
    alt_part.attach(MIMEText(html_body, "html", "utf-8"))
    msg.attach(alt_part)

    # Embed each screenshot as an inline CID image
    for index, attachment in enumerate(attachments, start=1):
        with attachment.open("rb") as fh:
            img = MIMEImage(fh.read(), _subtype="png")
        img.add_header("Content-ID", f"<screenshot_{index}>")
        img.add_header("Content-Disposition", "inline", filename=attachment.name)
        img.attach = None  # satisfy linters; already constructed
        msg.attach(img)

    context = None
    if settings.smtp_skip_verify:
        context = ssl._create_unverified_context()
    else:
        context = ssl.create_default_context()

    envelope_from = settings.smtp_envelope_from or settings.email_from
    logger.info("Sending email to %s with %s inline screenshot(s).", settings.email_to, len(attachments))
    
    if settings.smtp_use_ssl:
        smtp_client = smtplib.SMTP_SSL(settings.smtp_host, settings.smtp_port, timeout=60, context=context)
    else:
        smtp_client = smtplib.SMTP(settings.smtp_host, settings.smtp_port, timeout=60)

    with smtp_client as smtp:
        smtp.ehlo()
        if not settings.smtp_use_ssl and settings.smtp_use_tls:
            smtp.starttls(context=context)
            smtp.ehlo()
        if settings.smtp_username and settings.smtp_password:
            smtp.login(settings.smtp_username, settings.smtp_password)
        try:
            smtp.sendmail(envelope_from, settings.email_to, msg.as_string())
        except smtplib.SMTPDataError as error:
            response_text = ""
            if len(error.args) > 1 and isinstance(error.args[1], (bytes, bytearray)):
                response_text = error.args[1].decode("utf-8", errors="replace")
            elif len(error.args) > 1:
                response_text = str(error.args[1])

            if error.smtp_code == 550 and "5.7.60" in response_text:
                raise RuntimeError(
                    "SMTP rejected the sender with '5.7.60 Send As denied'. "
                    f"Authenticated SMTP user: {settings.smtp_username or '(anonymous)'} | "
                    f"EMAIL_FROM: {settings.email_from} | "
                    f"SMTP_ENVELOPE_FROM: {envelope_from}. "
                    "Set EMAIL_FROM and SMTP_ENVELOPE_FROM to a mailbox this SMTP account is allowed to send as, "
                    "or ask your mail admin to grant 'Send As' permission for that sender."
                ) from error
            raise


def main() -> int:
    settings = load_settings()
    ensure_directories(settings)
    logger = build_logger(settings)
    validate_settings(settings)
    validate_timezone(settings, logger)
    report_host = urlparse(settings.report_url).hostname
    if settings.auth_mode == "integrated" and is_ipv4_host(report_host):
        logger.warning(
            "REPORT_URL uses an IP address (%s). Integrated Windows authentication often works more reliably "
            "with a DNS hostname or FQDN because Kerberos/SPN matching can fail on raw IP addresses.",
            report_host,
        )
    if settings.auth_mode == "integrated":
        logger.info("Integrated auth allowlist: %s", build_auth_server_allowlist(settings))
    logger.info("Starting PBIRS capture job.")

    screenshots, errors = capture_report(settings, logger)
    send_email(settings, logger, screenshots, errors)

    if errors:
        logger.warning("Capture completed with %s tab error(s).", len(errors))
        return 1

    logger.info("Capture completed successfully.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        fallback_log_dir = Path(os.getenv("LOG_DIR", "logs"))
        fallback_log_dir.mkdir(parents=True, exist_ok=True)
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s | %(levelname)s | %(message)s",
            handlers=[
                logging.FileHandler(fallback_log_dir / "pbirs_capture.log", encoding="utf-8"),
                logging.StreamHandler(sys.stdout),
            ],
        )
        logging.exception("PBIRS capture job failed: %s", error)
        raise SystemExit(1)
