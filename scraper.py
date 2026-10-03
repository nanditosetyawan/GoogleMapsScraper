"""
scraper.py – Core Google Maps Reviews scraper using Playwright.

Stages:
  1.  URL validation & short-URL resolution
  2.  Browser launch with realistic context
  3.  Page load & CAPTCHA detection
  4.  Locate business page
  5.  Open reviews panel
  6.  Find scroll container
  7.  Auto-scroll
  8.  Extract review cards
  9.  Field extraction (name, rating, comment, date)
  10. Deduplication
  11. Checkpoint saving
  12. Export XLSX & CSV
"""

from __future__ import annotations

import csv
import json
import logging
import re
import time
from playwright_stealth import Stealth
from datetime import datetime
from pathlib import Path
from typing import Optional
from urllib.parse import urlparse

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from playwright.sync_api import (
    Browser,
    BrowserContext,
    Page,
    Playwright,
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError,
)

from config import Config

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------

def normalize_text(text: str) -> str:
    """Strip and collapse whitespace; preserve original casing."""
    if not text:
        return ""
    return re.sub(r"\s+", " ", text.strip())


def normalize_for_dedup(text: str) -> str:
    """Lowercase + normalize for deduplication key only."""
    return normalize_text(text).lower()


def normalize_rating(raw: str) -> Optional[int]:
    """
    Convert various rating representations to int 1-5.
    Examples:
      "5 stars" -> 5
      "4,0 von 5 Sternen" -> 4
      "3" -> 3
    """
    if not raw:
        return None
    match = re.search(r"(\d)[.,]?\d*\s*(?:star|stern|étoile|estrella|star|bintang)?", raw, re.IGNORECASE)
    if match:
        val = int(match.group(1))
        if 1 <= val <= 5:
            return val
    # Try plain digit
    digits = re.findall(r"\d", raw)
    if digits:
        val = int(digits[0])
        if 1 <= val <= 5:
            return val
    return None


def safe_get_text(locator, timeout: int = 1000) -> str:
    """Return inner text from locator instantly, or empty string on failure."""
    try:
        text = locator.evaluate("el => el.innerText", timeout=timeout)
        return normalize_text(text)
    except Exception:
        return ""


def safe_get_attribute(locator, attr: str, timeout: int = 3000) -> str:
    """Return attribute value from locator, or empty string on failure."""
    try:
        val = locator.get_attribute(attr, timeout=timeout)
        return normalize_text(val or "")
    except Exception:
        return ""


def extract_text_safe(page_or_locator, selector: str, timeout: int = 3000) -> str:
    """Try to get text from a CSS/xpath selector. Returns '' on failure."""
    try:
        el = page_or_locator.locator(selector).first
        return normalize_text(el.inner_text(timeout=timeout))
    except Exception:
        return ""


def extract_attribute_safe(page_or_locator, selector: str, attr: str, timeout: int = 3000) -> str:
    """Try to get attribute from a selector. Returns '' on failure."""
    try:
        el = page_or_locator.locator(selector).first
        val = el.get_attribute(attr, timeout=timeout)
        return normalize_text(val or "")
    except Exception:
        return ""


def save_debug_snapshot(page: Page, prefix: str, logs_dir: Path) -> None:
    """Save a screenshot and HTML snapshot for debugging."""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    try:
        png_path = logs_dir / f"{prefix}_{ts}.png"
        page.screenshot(path=str(png_path), full_page=False)
        logger.info("Screenshot saved: %s", png_path)
    except Exception as e:
        logger.warning("Failed to save screenshot: %s", e)

    try:
        html_path = logs_dir / f"{prefix}_{ts}.html"
        html_path.write_text(page.content(), encoding="utf-8")
        logger.info("HTML snapshot saved: %s", html_path)
    except Exception as e:
        logger.warning("Failed to save HTML: %s", e)


# ---------------------------------------------------------------------------
# URL helpers
# ---------------------------------------------------------------------------

VALID_GMAPS_HOSTS = {"maps.app.goo.gl", "www.google.com", "google.com", "maps.google.com"}


def is_valid_google_maps_url(url: str) -> bool:
    """Basic validation of Google Maps URL."""
    try:
        parsed = urlparse(url)
        netloc = parsed.netloc.lower()
        return any(netloc == h or netloc.endswith("." + h) for h in VALID_GMAPS_HOSTS)
    except Exception:
        return False


def is_short_url(url: str) -> bool:
    parsed = urlparse(url)
    return "maps.app.goo.gl" in parsed.netloc


# ---------------------------------------------------------------------------
# CAPTCHA detection
# ---------------------------------------------------------------------------

CAPTCHA_INDICATORS = [
    "recaptcha",
    "captcha",
    "unusual traffic",
    "lalu lintas tidak biasa",
    "verify you're not a robot",
    "verify you are human",
    "verifikasi",
    "not a robot",
    "i'm not a robot",
]


def is_captcha_page(page: Page) -> bool:
    try:
        content = page.content().lower()
        return any(ind in content for ind in CAPTCHA_INDICATORS)
    except Exception:
        return False


def wait_for_captcha_resolution(page: Page, timeout_ms: int = 300_000) -> None:
    """Block until the page is no longer a CAPTCHA page (or timeout)."""
    print("\n" + "=" * 60)
    print("  ⚠️  Google meminta verifikasi CAPTCHA!")
    print("  Silakan selesaikan verifikasi di browser.")
    print("  Program akan menunggu otomatis...")
    print("=" * 60 + "\n")

    deadline = time.time() + timeout_ms / 1000
    while time.time() < deadline:
        time.sleep(3)
        if not is_captcha_page(page):
            logger.info("CAPTCHA selesai – melanjutkan scraping.")
            return
    raise RuntimeError("Timeout menunggu penyelesaian CAPTCHA.")


# ---------------------------------------------------------------------------
# Stage 1: URL resolution
# ---------------------------------------------------------------------------

NETWORK_ERROR_KEYWORDS = [
    "wsarecv",
    "connection was forcibly closed",
    "connection reset",
    "net::err_connection",
    "stream reading error",
    "socket hang up",
    "econnreset",
    "network changed",
]


def is_network_error(exc: Exception) -> bool:
    """Return True if exception looks like a transient network error."""
    msg = str(exc).lower()
    return any(kw in msg for kw in NETWORK_ERROR_KEYWORDS)


def resolve_url(page: Page, url: str, cfg: Config, max_retries: int = 3) -> str:
    """Open URL in page and return the final resolved URL. Retries on network errors."""
    logger.info("Membuka URL: %s", url)

    for attempt in range(1, max_retries + 1):
        try:
            page.goto(url, timeout=cfg.navigation_timeout, wait_until="domcontentloaded")
            page.wait_for_load_state("networkidle", timeout=cfg.navigation_timeout)
            break  # success
        except PlaywrightTimeoutError:
            logger.warning("Timeout saat loading URL – melanjutkan dengan URL saat ini.")
            break
        except Exception as e:
            if is_network_error(e) and attempt < max_retries:
                wait_s = attempt * 5
                logger.warning(
                    "Network error (percobaan %d/%d): %s. Mencoba lagi dalam %ds...",
                    attempt, max_retries, e, wait_s,
                )
                time.sleep(wait_s)
                continue
            logger.warning("Error saat goto: %s – melanjutkan.", e)
            break

    if is_captcha_page(page):
        wait_for_captcha_resolution(page, cfg.captcha_wait_timeout)

    final_url = page.url
    logger.info("URL akhir: %s", final_url)
    return final_url


# ---------------------------------------------------------------------------
# Stage 3-4: Find business name and open reviews panel
# ---------------------------------------------------------------------------

REVIEW_BUTTON_SELECTORS = [
    # Aria-label patterns (various languages)
    "[aria-label*='review' i]",
    "[aria-label*='ulasan' i]",
    "[aria-label*='Reviews' i]",
    "[aria-label*='Ulasan' i]",
    # The rating button which opens reviews when clicked
    "button[jsaction*='moreReviews']",
    "button[aria-label*='bintang' i]",
    "button[aria-label*='star' i]",
    # Button containing review count
    "button[jsaction*='review']",
    "button[data-tab-index='1']",
    # Common role-based
    "[role='tab'][aria-label*='review' i]",
    "[role='tab'][aria-label*='ulasan' i]",
    # Text-based (last resort)
    "text=Reviews",
    "text=Ulasan",
]


def get_business_name(page: Page) -> str:
    """Try to extract business name for logging only."""
    candidates = [
        "h1",
        "[aria-label][role='heading']",
        ".fontHeadlineLarge",
        "[data-attrid='title']",
        ".x3AX1-LfntMc-header-title-ij8cu",
    ]
    for sel in candidates:
        name = extract_text_safe(page, sel)
        if name:
            return name
    return "(nama tidak ditemukan)"


def find_review_button(page: Page, cfg: Config):
    """
    Return a locator for the 'All reviews' / 'Ulasan' button.
    Tries multiple strategies and waits up to 15 seconds.
    """
    deadline = time.time() + 15.0
    while time.time() < deadline:
        # Strategy 1: role=tab with review label
        try:
            for label_pattern in ["review", "Reviews", "ulasan", "Ulasan"]:
                candidates = page.get_by_role("tab", name=re.compile(label_pattern, re.IGNORECASE))
                if candidates.count() > 0 and candidates.first.is_visible():
                    logger.debug("Review button found via role=tab with label '%s'", label_pattern)
                    return candidates.first
        except Exception:
            pass

        # Strategy 2: aria selectors
        for sel in REVIEW_BUTTON_SELECTORS:
            try:
                loc = page.locator(sel).first
                if loc.count() > 0 and loc.is_visible():
                    logger.debug("Review button found via selector: %s", sel)
                    return loc
            except Exception:
                continue

        # Strategy 3: text search
        for text in ["Reviews", "Ulasan", "review", "ulasan"]:
            try:
                loc = page.get_by_text(re.compile(rf"^\d+[\s,\.]+{text}", re.IGNORECASE)).first
                if loc.count() > 0 and loc.is_visible():
                    logger.debug("Review button found via text match: %s", text)
                    return loc
            except Exception:
                pass

        time.sleep(0.5)

    return None


def safe_click(locator, timeout: int = 10_000) -> bool:
    """Click a locator safely, returns True on success."""
    try:
        locator.scroll_into_view_if_needed(timeout=timeout)
        locator.click(timeout=timeout)
        return True
    except Exception as e:
        logger.debug("safe_click failed: %s", e)
        return False


def wait_for_any_selector(page: Page, selectors: list[str], timeout: int = 10_000):
    """Wait for the first matching selector; returns (selector, locator) or (None, None)."""
    deadline = time.time() + timeout / 1000
    while time.time() < deadline:
        for sel in selectors:
            try:
                loc = page.locator(sel).first
                if loc.count() > 0 and loc.is_visible(timeout=500):
                    return sel, loc
            except Exception:
                continue
        time.sleep(0.5)
    return None, None


def find_first_matching_locator(page: Page, selectors: list[str], timeout: int = 5000):
    """Return the first visible locator from a list of selectors."""
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            if loc.count() > 0 and loc.is_visible(timeout=timeout):
                return loc
        except Exception:
            continue
    return None


def open_reviews_panel(page: Page, cfg: Config) -> bool:
    """
    Click the review tab/button and wait for the review panel to appear.
    Returns True if successful.
    """
    logger.info("Mencari tombol ulasan...")

    btn = find_review_button(page, cfg)
    if btn is None:
        logger.error("Tombol ulasan tidak ditemukan.")
        save_debug_snapshot(page, "error_no_review_btn", cfg.logs_dir)
        return False

    if not safe_click(btn):
        logger.error("Gagal mengklik tombol ulasan.")
        save_debug_snapshot(page, "error_click_review_btn", cfg.logs_dir)
        return False

    logger.info("Membuka panel ulasan...")

    # Wait for reviews to appear
    REVIEW_PANEL_SELECTORS = [
        "[data-review-id]",
        "[jslog*='metadata:review']",
        ".jftiEf",
        "[aria-label*='review' i][role='listitem']",
        "[role='article']",
        ".bwb7ce",
        ".wiI7pd",
    ]
    sel, _ = wait_for_any_selector(page, REVIEW_PANEL_SELECTORS, timeout=15_000)
    if sel:
        logger.info("Panel ulasan terbuka. (detected via: %s)", sel)
        return True

    # Fallback: wait a bit and check for any review-like content
    time.sleep(3)
    content = page.content().lower()
    if "review" in content or "ulasan" in content:
        logger.info("Panel ulasan terbuka (fallback text check).")
        return True

    logger.warning("Panel ulasan mungkin tidak terbuka dengan benar.")
    save_debug_snapshot(page, "warn_reviews_panel", cfg.logs_dir)
    return False


# ---------------------------------------------------------------------------
# Stage 5: Find scroll container
# ---------------------------------------------------------------------------

SCROLL_CONTAINER_SELECTORS = [
    # Google Maps specific containers (order: most specific first)
    "div[aria-label*='review' i][role='feed']",
    "div[aria-label*='ulasan' i][role='feed']",
    "div[jslog*='reviews']",
    "div.m6QErb[aria-label]",
    "div.m6QErb.DxyBCb",
    "div.m6QErb.WNBkOb",
    "div.m6QErb",
    # Generic scrollable containers that contain reviews
    "div[role='feed']",
    "div[tabindex='-1'][aria-label*='review' i]",
]


def find_reviews_scroll_container(page: Page, cfg: Config):
    """
    Locate the scrollable container that holds the reviews.
    Returns a Playwright locator or None.
    """
    logger.info("Mencari review scroll container...")

    # Try known selectors
    for sel in SCROLL_CONTAINER_SELECTORS:
        try:
            loc = page.locator(sel).first
            if loc.count() == 0:
                continue
            # Check if it's scrollable via JS
            is_scrollable = page.evaluate(
                """(selector) => {
                    const el = document.querySelector(selector);
                    if (!el) return false;
                    return el.scrollHeight > el.clientHeight;
                }""",
                sel,
            )
            if is_scrollable:
                logger.info("Scroll container ditemukan: %s", sel)
                return loc
        except Exception:
            continue

    # Fallback: JS-based search for deepest scrollable div with review cards
    logger.info("Mencoba JS fallback untuk menemukan scroll container...")
    try:
        container_selector = page.evaluate(
            """() => {
                const reviewIndicators = ['[data-review-id]', '.jftiEf', '[aria-label*="review" i]', '.wiI7pd'];
                const allDivs = Array.from(document.querySelectorAll('div'));
                
                for (const div of allDivs) {
                    const style = window.getComputedStyle(div);
                    const isScrollable = (
                        (style.overflowY === 'auto' || style.overflowY === 'scroll') &&
                        div.scrollHeight > div.clientHeight + 50
                    );
                    if (!isScrollable) continue;
                    
                    // Check if it contains review cards
                    for (const indicator of reviewIndicators) {
                        if (div.querySelectorAll(indicator).length >= 2) {
                            // Return a unique path to this element
                            let path = '';
                            let el = div;
                            while (el && el !== document.body) {
                                let selector = el.tagName.toLowerCase();
                                if (el.id) {
                                    selector += '#' + el.id;
                                } else if (el.className && typeof el.className === 'string') {
                                    const classes = el.className.trim().split(/[\s]+/).slice(0, 2).join('.');
                                    if (classes) selector += '.' + classes;
                                }
                                path = selector + (path ? ' > ' + path : '');
                                el = el.parentElement;
                            }
                            return {found: true, classes: div.className, selector: indicator};
                        }
                    }
                }
                return {found: false};
            }"""
        )
        if container_selector and container_selector.get("found"):
            logger.info("JS fallback: container found with indicator: %s", container_selector.get("selector"))
    except Exception as e:
        logger.debug("JS fallback error: %s", e)

    # Last resort: try any div with overflow scroll/auto
    logger.info("Mencoba metode terakhir: mencari div scrollable manapun...")
    try:
        result = page.evaluate(
            """() => {
                const divs = Array.from(document.querySelectorAll('div'));
                const scrollable = divs.filter(d => {
                    const s = window.getComputedStyle(d);
                    return (s.overflowY === 'scroll' || s.overflowY === 'auto') 
                        && d.scrollHeight > d.clientHeight + 100
                        && d.clientHeight > 200;
                });
                // Return className of largest scrollable div
                if (scrollable.length === 0) return null;
                scrollable.sort((a, b) => b.scrollHeight - a.scrollHeight);
                return scrollable[0].className || null;
            }"""
        )
        if result:
            first_class = result.strip().split()[0] if result.strip() else None
            if first_class:
                sel = f"div.{first_class}"
                loc = page.locator(sel).first
                if loc.count() > 0:
                    logger.info("Last resort container found via class: %s", first_class)
                    return loc
    except Exception as e:
        logger.debug("Last resort container search error: %s", e)

    logger.error("Review scroll container tidak ditemukan!")
    save_debug_snapshot(page, "error_no_scroll_container", cfg.logs_dir)
    return None


# ---------------------------------------------------------------------------
# Stage 7-8: Review card detection and field extraction
# ---------------------------------------------------------------------------

REVIEW_CARD_SELECTORS = [
    "[data-review-id]",
    "div[jslog*='metadata:review']",
    ".jftiEf",
    "div[role='article']",
    ".wiI7pd",          # review text container parent
    "[aria-label*='star' i][role='img']",   # rating presence indicator
]

# Selectors for sub-fields within a review card
NAME_SELECTORS = [
    ".d4r55",           # reviewer name common class
    "[aria-label]",
    ".kvMYJc",
    "button[data-href*='/maps/contrib/']",
    "a[href*='/maps/contrib/']",
    "span[class*='name' i]",
    ".WNxzHc a",
    ".X43Kjb",
]

RATING_SELECTORS = [
    "span[aria-label*='star' i]",
    "[role='img'][aria-label*='star' i]",
    "span[role='img']",
    ".kvMYJc",
    "div[aria-label*='star' i]",
    "span[class*='rating' i]",
]

COMMENT_SELECTORS = [
    ".MyEned span.wiI7pd",
    "span.wiI7pd",
    ".wiI7pd",
    "[data-expandable-section]",
    "span[jslog*='review']",
    ".review-full-text",
    "[class*='review' i] span",
]

DATE_SELECTORS = [
    "span.rsqaWe",
    "span[class*='date' i]",
    ".DU9Pgb span",
    "span[aria-label*='ago' i]",
    ".y3Ibjb",
    "[class*='time' i]",
    "span[jslog*='timestamp']",
]


def expand_review_text(card, page: Page) -> None:
    """Click 'More'/'Selengkapnya' button if present in the card instantly."""
    MORE_SELECTORS = [
        "button[aria-label*='more' i]",
        "button.w8nwRe",
        "button[jsaction*='expand' i]",
        "button.review-more-link",
        "span.review-more-link",
    ]
    for sel in MORE_SELECTORS:
        try:
            btn = card.locator(sel).first
            if btn.count() > 0:
                btn.evaluate("el => { if (el.offsetWidth > 0 && el.offsetHeight > 0) el.click(); }", timeout=500)
                return
        except Exception:
            continue

    # Try text-based
    for text in ["More", "Selengkapnya", "Lainnya"]:
        try:
            btn = card.get_by_text(text).first
            if btn.count() > 0:
                btn.evaluate("el => { if (el.offsetWidth > 0 && el.offsetHeight > 0) el.click(); }", timeout=500)
                return
        except Exception:
            continue


def extract_reviewer_name(card) -> str:
    """Extract reviewer name from a review card."""
    # Priority: button or anchor with /maps/contrib/
    for sel in ["button[data-href*='/maps/contrib/']", "a[href*='/maps/contrib/']"]:
        try:
            loc = card.locator(sel).first
            if loc.count() > 0:
                text = safe_get_text(loc)
                if text:
                    return text
                # Try aria-label
                label = safe_get_attribute(loc, "aria-label")
                if label:
                    return label
        except Exception:
            continue

    # Try class-based selectors
    for sel in [".d4r55", ".kvMYJc", ".WNxzHc", ".X43Kjb", ".fontBodyMedium"]:
        try:
            loc = card.locator(sel).first
            if loc.count() > 0:
                text = safe_get_text(loc)
                if text and len(text) > 1:
                    return text
        except Exception:
            continue

    # Try any aria-label on reviewer container
    for sel in ["[data-review-id]", "[data-href*='maps']"]:
        try:
            label = safe_get_attribute(card.locator(sel).first, "aria-label")
            if label:
                # aria-label might be "reviewer's review: ..."
                match = re.match(r"^(.+?)(?:'s| )", label)
                if match:
                    return match.group(1)
        except Exception:
            continue

    return "(anonim)"


def extract_rating(card) -> Optional[int]:
    """Extract rating (1-5) from a review card."""
    for sel in RATING_SELECTORS:
        try:
            loc = card.locator(sel).first
            if loc.count() == 0:
                continue
            # Check aria-label first
            label = safe_get_attribute(loc, "aria-label")
            if label:
                rating = normalize_rating(label)
                if rating:
                    return rating
            # Check title attribute
            title = safe_get_attribute(loc, "title")
            if title:
                rating = normalize_rating(title)
                if rating:
                    return rating
        except Exception:
            continue

    # JS approach: look for aria-label with star count
    try:
        rating_text = card.evaluate(
            """(el) => {
                const imgs = el.querySelectorAll('[role="img"]');
                for (const img of imgs) {
                    const label = img.getAttribute('aria-label');
                    if (label && /star/i.test(label)) return label;
                }
                return null;
            }"""
        )
        if rating_text:
            return normalize_rating(rating_text)
    except Exception:
        pass

    return None


def extract_comment(card) -> str:
    """Extract review comment text from a review card."""
    for sel in COMMENT_SELECTORS:
        try:
            loc = card.locator(sel).first
            if loc.count() == 0:
                continue
            text = safe_get_text(loc)
            if text and len(text) > 2:
                return text
        except Exception:
            continue

    # JS fallback: get the longest text block
    try:
        text = card.evaluate(
            """(el) => {
                const spans = el.querySelectorAll('span');
                let longest = '';
                for (const span of spans) {
                    const t = span.innerText || '';
                    if (t.length > longest.length && !t.includes('star')) {
                        longest = t;
                    }
                }
                return longest;
            }"""
        )
        if text and len(text) > 5:
            return normalize_text(text)
    except Exception:
        pass

    return ""


def extract_date(card) -> str:
    """Extract review date as shown in Google Maps."""
    for sel in DATE_SELECTORS:
        try:
            loc = card.locator(sel).first
            if loc.count() == 0:
                continue
            # Check aria-label for precise date
            label = safe_get_attribute(loc, "aria-label")
            if label:
                return label
            # Check title attribute
            title = safe_get_attribute(loc, "title")
            if title:
                return title
            # Use visible text
            text = safe_get_text(loc)
            if text:
                return text
        except Exception:
            continue

    # JS fallback: look for time-like patterns
    try:
        date_text = card.evaluate(
            """(el) => {
                const patterns = [
                    /\\d+\\s+(day|week|month|year|hari|minggu|bulan|tahun)s?\\s+ago/i,
                    /\\d{1,2}\\s+\\w+\\s+\\d{4}/,
                ];
                const allText = el.innerText;
                for (const p of patterns) {
                    const m = allText.match(p);
                    if (m) return m[0];
                }
                return null;
            }"""
        )
        if date_text:
            return normalize_text(date_text)
    except Exception:
        pass

    return ""


def extract_review_data(card, page: Page) -> Optional[dict]:
    """
    Extract all fields from a single review card.
    Returns dict with keys: nama_pengulas, rating, komentar, tanggal
    Returns None if essential data (rating) is missing.
    """
    # Expand "More" if needed
    expand_review_text(card, page)

    nama_pengulas = extract_reviewer_name(card)
    rating = extract_rating(card)
    komentar = extract_comment(card)
    tanggal = extract_date(card)

    if rating is None:
        logger.debug("Rating tidak ditemukan untuk review card – dilewati.")
        return None

    return {
        "nama_pengulas": normalize_text(nama_pengulas),
        "rating": rating,
        "komentar": normalize_text(komentar),
        "tanggal": normalize_text(tanggal),
    }


# ---------------------------------------------------------------------------
# Stage 7: Find review cards
# ---------------------------------------------------------------------------

def find_review_cards(container, page: Page) -> list:
    """Return a list of Playwright locators for individual review cards."""
    # Try each selector
    for sel in REVIEW_CARD_SELECTORS:
        try:
            cards = container.locator(sel)
            count = cards.count()
            if count >= 1:
                logger.debug("Review cards found via '%s': %d cards", sel, count)
                return [cards.nth(i) for i in range(count)]
        except Exception:
            continue

    # Fallback: JS to detect review card structure
    logger.warning("Trying JS fallback for review cards...")
    try:
        count = container.evaluate(
            """(el) => {
                const candidates = [
                    '[data-review-id]',
                    'div[jslog*="metadata:review"]',
                    '.jftiEf',
                    '[role="article"]',
                ];
                for (const sel of candidates) {
                    const found = el.querySelectorAll(sel);
                    if (found.length > 0) return found.length;
                }
                return 0;
            }"""
        )
        if count > 0:
            # Use the working selector
            for sel in REVIEW_CARD_SELECTORS:
                try:
                    cards = page.locator(sel)
                    if cards.count() >= 1:
                        cnt = cards.count()
                        return [cards.nth(i) for i in range(cnt)]
                except Exception:
                    continue
    except Exception as e:
        logger.debug("JS fallback for cards failed: %s", e)

    return []


# ---------------------------------------------------------------------------
# Stage 9: Deduplication
# ---------------------------------------------------------------------------

def make_dedup_key(review: dict) -> tuple:
    return (
        normalize_for_dedup(review.get("nama_pengulas", "")),
        review.get("rating", 0),
        normalize_for_dedup(review.get("komentar", "")[:200]),
        normalize_for_dedup(review.get("tanggal", "")),
    )


# ---------------------------------------------------------------------------
# Stage 10: Checkpoint
# ---------------------------------------------------------------------------

def load_checkpoint(cfg, url):
    return [], set()


def save_checkpoint(cfg, url, reviews):
    pass


# ---------------------------------------------------------------------------
# Stage 6: Auto-scroll
# ---------------------------------------------------------------------------

def scroll_and_collect(
    page: Page,
    container,
    cfg: Config,
    url: str,
    existing_reviews: list[dict],
    existing_keys: set,
) -> list[dict]:
    """
    Scroll the review container, collect and deduplicate reviews.
    Saves checkpoints periodically.
    """
    reviews = list(existing_reviews)
    dedup_keys = set(existing_keys)
    no_new_count = 0
    scroll_iteration = 0
    last_checkpoint_count = len(reviews)
    last_card_count = 0

    logger.info("Memulai scrolling otomatis...")

    while True:
        scroll_iteration += 1

        # 1. Klik semua tombol 'Selengkapnya' (More) secara massal via JS untuk kartu-kartu baru
        newly_added = 0
        container.evaluate("""(container, startIdx) => {
            const cards = Array.from(container.querySelectorAll('.jftiEf, [data-review-id]')).slice(startIdx);
            for (const card of cards) {
                const btns = card.querySelectorAll('button[aria-label*="more" i], button.w8nwRe, button.review-more-link, span.review-more-link');
                for (const btn of btns) {
                    if (btn.offsetWidth > 0 && btn.offsetHeight > 0) btn.click();
                }
                const textBtns = Array.from(card.querySelectorAll('button')).filter(b => /more|selengkapnya|lainnya/i.test(b.innerText));
                for (const btn of textBtns) {
                    if (btn.offsetWidth > 0 && btn.offsetHeight > 0) btn.click();
                }
            }
        }""", last_card_count)

        # Beri waktu sejenak agar teks ulasan yang panjang terbuka
        time.sleep(0.3)

        # 2. Ekstrak data semua kartu baru secara massal via JS (Kecepatan Ultra / O(1) CDP call)
        new_data = container.evaluate("""(container, startIdx) => {
            const cards = Array.from(container.querySelectorAll('.jftiEf, [data-review-id]')).slice(startIdx);
            const results = [];
            
            for (const card of cards) {
                let nama = "Anonim";
                // Gunakan class spesifik yang biasanya memuat nama reviewer di Google Maps
                const nameEl = card.querySelector('button[data-href*="/maps/contrib/"] div.d4r55, a[href*="/maps/contrib/"] div, .d4r55, .WNxzHc a, .X43Kjb');
                if (nameEl) nama = nameEl.innerText || nameEl.textContent;
                else {
                    const btn = card.querySelector('button[data-href*="/maps/contrib/"]');
                    if (btn) {
                        const aria = btn.getAttribute('aria-label');
                        if (aria) nama = aria;
                    }
                }
                
                let rating = null;
                const starEls = card.querySelectorAll('[role="img"]');
                for (const el of starEls) {
                    const label = el.getAttribute('aria-label');
                    if (label && /star|bintang/i.test(label)) {
                        const m = label.match(/(\\d)[.,]?\\d*/);
                        if (m) { rating = parseInt(m[1]); break; }
                    }
                }
                if (!rating) {
                    const starSpan = card.querySelector('.kvMYJc');
                    if (starSpan) {
                        const label = starSpan.getAttribute('aria-label');
                        if (label) {
                            const m = label.match(/(\\d)[.,]?\\d*/);
                            if (m) rating = parseInt(m[1]);
                        }
                    }
                }
                
                let komentar = "";
                const commentEl = card.querySelector('.MyEned span.wiI7pd, span.wiI7pd, .wiI7pd');
                if (commentEl) komentar = commentEl.innerText || commentEl.textContent;
                
                let tanggal = "";
                // Hindari selector yang salah ambil icon bintang (DU9Pgb)
                const dateEl = card.querySelector('span.rsqaWe, span.xILg5e, span[class*="date" i], span[aria-label*="ago" i], .y3Ibjb, span[jslog*="timestamp"]');
                if (dateEl) tanggal = dateEl.innerText || dateEl.textContent;
                
                // Jika tanggal memuat simbol bintang (\\ue838), abaikan
                if (tanggal.includes('\\ue838')) tanggal = "";
                
                results.push({
                    nama_pengulas: (nama || "").replace(/\\s+/g, ' ').trim(),
                    rating: rating,
                    komentar: (komentar || "").replace(/\\s+/g, ' ').trim(),
                    tanggal: (tanggal || "").replace(/\\s+/g, ' ').trim()
                });
            }
            return { totalCards: document.querySelectorAll('.jftiEf, [data-review-id]').length, results: results };
        }""", last_card_count)

        current_card_count = new_data["totalCards"]
        
        for review in new_data["results"]:
            if not review.get("rating"):
                continue
            key = make_dedup_key(review)
            if key not in dedup_keys:
                dedup_keys.add(key)
                reviews.append(review)
                newly_added += 1

        if newly_added > 0:
            no_new_count = 0
            logger.info("Review unik terkumpul: %d", len(reviews))
        else:
            no_new_count += 1
            logger.info(
                "Tidak ada review baru (percobaan %d/%d)",
                no_new_count,
                cfg.max_scroll_attempts_without_new_reviews,
            )

        # Perbarui tracker jumlah kartu
        last_card_count = current_card_count

        # Check max_reviews limit
        if cfg.max_reviews > 0 and len(reviews) >= cfg.max_reviews:
            logger.info("Batas max_reviews=%d tercapai.", cfg.max_reviews)
            break

        # Stop condition
        if no_new_count >= cfg.max_scroll_attempts_without_new_reviews:
            logger.info("Tidak ada review baru – scraping dianggap selesai.")
            break

        # Save checkpoint periodically
        if len(reviews) - last_checkpoint_count >= cfg.checkpoint_batch_size:
            save_checkpoint(cfg, url, reviews)
            last_checkpoint_count = len(reviews)

        # Scroll down – dengan retry pada network error
        scroll_ok = False
        logger.info("Melakukan scroll ke bawah... (iterasi %d)", scroll_iteration)
        for scroll_try in range(3):
            try:
                container.evaluate("(el) => { el.scrollTop = el.scrollHeight; }")
                scroll_ok = True
                break
            except Exception as se:
                if is_network_error(se):
                    logger.warning(
                        "Network error saat scroll (percobaan %d/3): %s. Menunggu 5 detik...",
                        scroll_try + 1, se,
                    )
                    time.sleep(5)
                else:
                    # Non-network error, try alternate scroll methods
                    try:
                        container.press("End")
                        scroll_ok = True
                    except Exception:
                        try:
                            page.keyboard.press("End")
                            scroll_ok = True
                        except Exception:
                            pass
                    break

        if not scroll_ok:
            logger.warning("Tidak dapat melakukan scroll – melanjutkan tanpa scroll.")

        # Dynamic pause: shorter if we're finding new reviews, longer if not
        pause = cfg.scroll_pause if newly_added > 0 else min(cfg.scroll_pause * 1.5, 5.0)
        time.sleep(pause)

    # Final checkpoint
    save_checkpoint(cfg, url, reviews)
    return reviews


# ---------------------------------------------------------------------------
# Stage 11-12: Export
# ---------------------------------------------------------------------------

REQUIRED_COLUMNS = ["nama_pengulas", "rating", "komentar", "tanggal"]


def export_xlsx(reviews: list[dict], cfg: Config) -> Path:
    """Export reviews to formatted Excel file."""
    df = pd.DataFrame(reviews, columns=REQUIRED_COLUMNS)
    df["rating"] = pd.to_numeric(df["rating"], errors="coerce").astype("Int64")

    xlsx_path = cfg.output_xlsx
    df.to_excel(xlsx_path, index=False, sheet_name="Reviews")

    # Apply formatting with openpyxl
    wb = load_workbook(xlsx_path)
    ws = wb["Reviews"]

    # Header formatting
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=12)

    for col_idx, col_name in enumerate(REQUIRED_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Freeze top row
    ws.freeze_panes = "A2"

    # Auto-filter
    ws.auto_filter.ref = ws.dimensions

    # Column widths & wrap text for data
    col_widths = {
        "nama_pengulas": 25,
        "rating": 10,
        "komentar": 60,
        "tanggal": 20,
    }

    for col_idx, col_name in enumerate(REQUIRED_COLUMNS, start=1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = col_widths.get(col_name, 20)

        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.alignment = Alignment(
                wrap_text=(col_name == "komentar"),
                vertical="top",
            )

    # Row height for header
    ws.row_dimensions[1].height = 25

    wb.save(xlsx_path)
    logger.info("Excel disimpan: %s", xlsx_path)
    return xlsx_path


def export_csv(reviews: list[dict], cfg: Config) -> Path:
    """Export reviews to UTF-8-SIG CSV (Excel compatible)."""
    csv_path = cfg.output_csv
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=REQUIRED_COLUMNS)
        writer.writeheader()
        for review in reviews:
            writer.writerow({k: review.get(k, "") for k in REQUIRED_COLUMNS})
    logger.info("CSV disimpan: %s", csv_path)
    return csv_path


# ---------------------------------------------------------------------------
# Stage 18: Summary
# ---------------------------------------------------------------------------

def get_total_review_count_from_page(page: Page) -> Optional[int]:
    """Try to extract the total review count shown by Google Maps."""
    patterns = [
        r"([\d,\.]+)\s+review",
        r"([\d,\.]+)\s+ulasan",
        r"([\d,\.]+)\s+Rezension",
    ]
    try:
        content = page.content()
        for pat in patterns:
            match = re.search(pat, content, re.IGNORECASE)
            if match:
                num_str = match.group(1).replace(",", "").replace(".", "")
                return int(num_str)
    except Exception:
        pass
    return None


def print_summary(url: str, reviews: list[dict], xlsx_path: Path, csv_path: Path, gmaps_total: Optional[int]) -> None:
    """Print final summary to terminal."""
    sep = "=" * 44
    print(f"\n{sep}")
    print("SCRAPING SELESAI")
    print(sep)
    print(f"\nURL:\n{url}\n")
    print(f"Total review unik:\n{len(reviews)}\n")
    print(f"Output:\n{xlsx_path}\n{csv_path}")
    print(f"\n{sep}\n")

    if gmaps_total and gmaps_total > len(reviews):
        print(f"[WARNING]")
        print(f"Google Maps menampilkan sekitar {gmaps_total} review,")
        print(f"tetapi scraper hanya mendapatkan {len(reviews)} review unik.")
        print("Kemungkinan:")
        print("  - pagination/dynamic loading")
        print("  - review tidak seluruhnya dimuat")
        print("  - Google membatasi hasil")
        print(f"  - struktur DOM berubah.\n")


# ---------------------------------------------------------------------------
# Main scraper entry point
# ---------------------------------------------------------------------------

class GoogleMapsReviewScraper:
    def __init__(self, cfg: Config):
        self.cfg = cfg

    def run(self) -> list[dict]:
        """Execute the full scraping pipeline. Returns list of review dicts."""
        cfg = self.cfg

        # Validate URL
        url = cfg.google_maps_url.strip()
        if not url:
            url = input("Masukkan URL Google Maps: ").strip()
            if not url:
                raise ValueError("URL tidak boleh kosong.")

        if not is_valid_google_maps_url(url):
            raise ValueError(f"URL tidak valid: {url}")

        with sync_playwright() as playwright:
            return self._run_with_playwright(playwright, url)

    def _run_with_playwright(self, playwright: Playwright, url: str) -> list[dict]:
        cfg = self.cfg
        
        # Use actual Google Chrome to bypass Google's "Insecure Browser" login block
        user_data_dir = cfg.output_dir / "chrome_profile"
        user_data_dir.mkdir(exist_ok=True)
        
        context: BrowserContext = playwright.chromium.launch_persistent_context(
            user_data_dir=str(user_data_dir),
            channel="chrome",
            headless=cfg.headless,
            args=["--disable-blink-features=AutomationControlled"],
            viewport={"width": cfg.viewport_width, "height": cfg.viewport_height},
            locale=cfg.locale,
            timezone_id=cfg.timezone,
            user_agent=cfg.user_agent,
            java_script_enabled=True,
        )

        # Remove automation detection signals
        context.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
        """)

        # When using persistent context, it already creates a default page
        page: Page = context.pages[0] if context.pages else context.new_page()
        Stealth().apply_stealth_sync(page)
        
        page.set_default_timeout(cfg.action_timeout)
        page.set_default_navigation_timeout(cfg.navigation_timeout)

        try:
            return self._scrape(page, url)
        except Exception as e:
            logger.error("Fatal error: %s", e, exc_info=True)
            save_debug_snapshot(page, "fatal_error", cfg.logs_dir)
            raise
        finally:
            # Keep browser open briefly then close
            logger.info("Browser akan ditutup dalam 3 detik...")
            time.sleep(3)
            context.close()

    def _scrape(self, page: Page, url: str) -> list[dict]:
        cfg = self.cfg

        # Stage 1: Resolve URL
        logger.info("Membuka Google Maps...")
        final_url = resolve_url(page, url, cfg)

        # Wait for Google Maps to stabilize
        time.sleep(3)

        if is_captcha_page(page):
            wait_for_captcha_resolution(page, cfg.captcha_wait_timeout)

        # Stage 3: Business page
        business_name = get_business_name(page)
        logger.info("Menemukan halaman bisnis: %s", business_name)

        # Load checkpoint
        reviews, dedup_keys = load_checkpoint(cfg, final_url)

        # Stage 4: Open reviews panel
        # Retry membuka panel ulasan jika gagal
        panel_opened = False
        for panel_attempt in range(1, 4):
            if open_reviews_panel(page, cfg):
                panel_opened = True
                break
            logger.warning(
                "Percobaan membuka panel ulasan %d/3 gagal. Mencoba lagi...", panel_attempt
            )
            time.sleep(3)

        if not panel_opened:
            logger.warning("Gagal menemukan/mengklik tab ulasan. Mencoba lanjut mencari scroll container (mungkin tab disembunyikan Google).")

        time.sleep(2)
        
        # Coba klik tombol "Ulasan lainnya" jika kita berada dalam mode preview pencarian
        try:
            logger.info("Mencoba mencari tombol 'Ulasan lainnya'...")
            js_code = """() => {
                const els = Array.from(document.querySelectorAll('span, button, div, a'));
                for (let el of els) {
                    const text = (el.textContent || '').toLowerCase();
                    const aria = (el.getAttribute('aria-label') || '').toLowerCase();
                    if (text.includes('ulasan lainnya') || text.includes('more reviews') || 
                        aria.includes('ulasan lainnya') || aria.includes('more reviews')) {
                        el.click();
                        return true;
                    }
                }
                return false;
            }"""
            clicked = page.evaluate(js_code)
            if clicked:
                logger.info("Berhasil mengklik 'Ulasan lainnya' via JavaScript!")
                time.sleep(3)
        except Exception as e:
            logger.warning("Gagal mengklik Ulasan lainnya: %s", e)

        # Stage 5: Find scroll container
        # Coba beberapa kali karena panel mungkin masih loading
        container = None
        for container_attempt in range(1, 4):
            container = find_reviews_scroll_container(page, cfg)
            if container is not None:
                break
            logger.warning(
                "Scroll container belum ditemukan (percobaan %d/3) – menunggu...", container_attempt
            )
            time.sleep(3)

        if container is None:
            raise RuntimeError("Review scroll container tidak ditemukan. Cek logs/ untuk screenshot.")

        logger.info("Review container ditemukan.")

        # Get total count shown by Google Maps (for summary warning)
        gmaps_total = get_total_review_count_from_page(page)
        if gmaps_total:
            logger.info("Google Maps menampilkan sekitar %d review total.", gmaps_total)

        # Stages 6-10: Scroll, extract, deduplicate, checkpoint
        reviews = scroll_and_collect(page, container, cfg, final_url, reviews, dedup_keys)

        logger.info("Scraping selesai. Total review: %d", len(reviews))

        if not reviews:
            logger.warning("Tidak ada review yang berhasil diambil.")
            return reviews

        # Stage 11-12: Export
        xlsx_path = export_xlsx(reviews, cfg)
        csv_path = export_csv(reviews, cfg)

        logger.info("Excel: %s", xlsx_path)
        logger.info("CSV: %s", csv_path)

        # Stage 18: Summary
        print_summary(final_url, reviews, xlsx_path, csv_path, gmaps_total)

        return reviews
