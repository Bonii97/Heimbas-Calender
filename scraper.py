import os
import sys
import re
import hashlib
from datetime import datetime, timedelta, timezone
from typing import List, Optional, Tuple, Dict, Any

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from bs4 import BeautifulSoup
from icalendar import Calendar, Event
from zoneinfo import ZoneInfo


BERLIN_TZ = ZoneInfo("Europe/Berlin")


def debug(msg: str) -> None:
    """Lightweight debug logger to stderr."""
    print(f"[scraper] {msg}", file=sys.stderr)


def get_env_credentials() -> Tuple[str, str]:
    """Read credentials from environment variables.

    Returns a tuple of (username, password). Exits if missing.
    """
    user = os.environ.get("HEIMBAS_USER", "").strip()
    pw = os.environ.get("HEIMBAS_PASS", "").strip()
    if not user or not pw:
        print("Fehler: HEIMBAS_USER und/oder HEIMBAS_PASS sind nicht gesetzt.", file=sys.stderr)
        sys.exit(2)
    return user, pw


def try_fill(page, selectors: List[str], value: str) -> bool:
    """Try multiple selectors until one works. Returns True on success."""
    for sel in selectors:
        try:
            locator = page.locator(sel)
            if locator.count() > 0:
                locator.first.fill(value)
                return True
        except Exception:
            continue
    return False


def try_click(page, selectors_or_text: List[str]) -> bool:
    """Try to click either css/xpath selectors or buttons/links by text."""
    for sel in selectors_or_text:
        try:
            if sel.startswith("css="):
                page.locator(sel[len("css="):]).first.click()
                return True
            if sel.startswith("xpath="):
                page.locator(sel).first.click()
                return True
            # Fallback: try get_by_role or get_by_text
            btn = page.get_by_role("button", name=re.compile(sel, re.I))
            if btn.count() > 0:
                btn.first.click()
                return True
            link = page.get_by_role("link", name=re.compile(sel, re.I))
            if link.count() > 0:
                link.first.click()
                return True
            # Try visible text anywhere
            el = page.get_by_text(re.compile(sel, re.I))
            if el.count() > 0:
                el.first.click()
                return True
        except Exception:
            continue
    return False


def login_and_get_einsatz_vorschau_html(base_url: str, username: str, password: str) -> str:
    """Log in via Playwright (Chromium, headless) and navigate to 'Einsatz-Vorschau'.

    Returns the page HTML content containing the table. If the expected table
    cannot be found within the page, raises RuntimeError.
    """
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()

        debug("Rufe Login-Seite auf…")
        page.goto(base_url, wait_until="domcontentloaded", timeout=60_000)

        # Try to locate username/password fields in a robust way
        user_selectors = [
            'input[name="username"]',
            'input[id*="user" i]',
            'input[placeholder*="utzer" i]',  # Benutzername/Username
            'input[type="email"]',
            'input[type="text"]',
        ]
        pass_selectors = [
            'input[name="password"]',
            'input[id*="pass" i]',
            'input[placeholder*="ass" i]',
            'input[type="password"]',
        ]

        # Fill username/password if present on this page
        try:
            page.wait_for_timeout(1000)
            filled_user = try_fill(page, user_selectors, username)
            filled_pass = try_fill(page, pass_selectors, password)

            if filled_user and filled_pass:
                debug("Klicke auf Anmelden…")
                clicked = try_click(page, [
                    "Anmelden", "Login", "Einloggen", "Anmeldung",
                    'css=button[type="submit"]',
                    'css=button:has-text("Anmelden")',
                ])
                if not clicked:
                    # try pressing Enter in password field
                    page.locator(pass_selectors[0]).press("Enter")
            else:
                debug("Kein klassisches Login-Formular gefunden; ggf. SSO/Weiterleitung…")

            # Wait for navigation post-login
            page.wait_for_load_state("networkidle", timeout=60_000)
        except PlaywrightTimeoutError:
            pass

        # Try to navigate/click to 'Einsatz-Vorschau'
        debug("Versuche zur Seite 'Einsatz-Vorschau' zu wechseln…")
        # Click any obvious navigation element
        try_click(page, [
            "Einsatz-Vorschau", "Einsatz Vorschau", "Vorschau",
        ])

        # As a fallback, try to open likely paths within the app
        try:
            # Some apps accept direct navigation within the same auth context
            page.goto(base_url.rstrip('/') + "/einsatz-vorschau", wait_until="domcontentloaded", timeout=30_000)
        except Exception:
            pass

        # Wait for possible table to appear
        try:
            page.wait_for_timeout(1500)
            # Heuristic wait: look for a table element with headers
            page.wait_for_selector("table", timeout=20_000)
        except PlaywrightTimeoutError:
            pass

        html = page.content()
        context.close()
        browser.close()

    # Basic validation: ensure table likely exists
    if not contains_einsatz_table(html):
        with open("lastpage.html", "w", encoding="utf-8") as f:
            f.write(html)
        raise RuntimeError(
            "Konnte keine Einsatz-Tabelle finden. Die zuletzt geladene Seite wurde als 'lastpage.html' gespeichert."
        )
    return html


def contains_einsatz_table(html: str) -> bool:
    """Heuristically determine whether HTML contains the desired table."""
    soup = BeautifulSoup(html, "lxml")
    for table in soup.find_all("table"):
        header_text = " ".join(th.get_text(strip=True) for th in table.find_all(["th", "td"]))
        header_text_lower = header_text.lower()
        if any(keyword in header_text_lower for keyword in [
            "datum", "einsatz", "uhrzeit", "von", "bis", "beschreibung", "adresse"
        ]):
            return True
    return False


def parse_table_entries(html: str) -> List[Dict[str, Any]]:
    """Parse the HTML table and extract entries with date/time/description/address.

    Returns a list of dicts with keys: date, start_time, end_time, description, address.
    """
    soup = BeautifulSoup(html, "lxml")
    tables = soup.find_all("table")
    if not tables:
        raise RuntimeError("Keine Tabelle im HTML gefunden.")

    # Select the first plausible table by checking headers
    chosen = None
    for table in tables:
        header_text = " ".join(th.get_text(strip=True) for th in table.find_all(["th", "td"]))
        header_text_lower = header_text.lower()
        if any(k in header_text_lower for k in ["datum", "einsatz", "uhrzeit", "beschreibung", "adresse", "von", "bis"]):
            chosen = table
            break

    if chosen is None:
        raise RuntimeError("Keine passende Einsatz-Tabelle erkannt.")

    entries: List[Dict[str, Any]] = []
    rows = chosen.find_all("tr")
    for row in rows:
        cells = [c.get_text("\n", strip=True) for c in row.find_all(["td", "th"])]
        if len(cells) < 2:
            continue

        combined = " | ".join(cells)
        # Extract date and times from the combined row text
        date_str = extract_date(combined)
        start_str, end_str = extract_time_range(combined)

        # Description: try to find the longest or most descriptive cell
        description = infer_description(cells)
        address = infer_address(description, cells)

        if not date_str or not start_str:
            # Not enough information to build an event; skip header or invalid rows
            continue

        entries.append({
            "date": date_str,
            "start_time": start_str,
            "end_time": end_str,
            "description": description,
            "address": address,
        })

    if not entries:
        raise RuntimeError("Die Tabelle enthält keine auswertbaren Einsatz-Zeilen.")

    return entries


def extract_date(text: str) -> Optional[str]:
    """Extract German-style date (dd.mm.yyyy or dd.mm.yy) from text."""
    m = re.search(r"(\d{1,2}\.\d{1,2}\.(?:\d{2}|\d{4}))", text)
    if m:
        return m.group(1)
    return None


def extract_time_range(text: str) -> Tuple[Optional[str], Optional[str]]:
    """Extract time range like '08:00 - 09:30' or 'von 8:00 bis 9:30'."""
    # Pattern with dash
    m = re.search(r"(\d{1,2}:\d{2})\s*[–\-]\s*(\d{1,2}:\d{2})", text)
    if m:
        return m.group(1), m.group(2)
    # Pattern with 'von ... bis ...'
    m = re.search(r"von\s*(\d{1,2}:\d{2})\s*bis\s*(\d{1,2}:\d{2})", text, re.I)
    if m:
        return m.group(1), m.group(2)
    # Single time (fallback)
    m = re.search(r"(\d{1,2}:\d{2})", text)
    if m:
        return m.group(1), None
    return None, None


def infer_description(cells: List[str]) -> str:
    """Heuristic to derive the most descriptive text from row cells."""
    # Prefer the cell with the most characters (likely description)
    description = max(cells, key=lambda s: len(s))
    return description.strip()


def infer_address(description: str, cells: List[str]) -> Optional[str]:
    """Try to extract an address from the description or other cells.

    Heuristics: look for lines with a German ZIP code, lines prefixed by 'Adresse', or
    the last line if it resembles an address.
    """
    # Inspect description lines
    cand_lines = [l.strip() for l in description.split("\n") if l.strip()]
    # Search for "Adresse: ..."
    for line in cand_lines:
        m = re.search(r"adresse\s*[:\-]\s*(.+)$", line, re.I)
        if m:
            return m.group(1).strip()
    # Search for a postal code line
    for line in cand_lines:
        if re.search(r"\b\d{5}\b", line):
            return line
    # If none found, scan all cells for a likely address
    for c in cells:
        if re.search(r"\b\d{5}\b", c):
            return c.strip()
    # Fallback: last line if moderately long
    if cand_lines:
        last = cand_lines[-1]
        if len(last) > 10:
            return last
    return None


def parse_german_date(date_str: str) -> datetime:
    """Parse dd.mm.yyyy or dd.mm.yy to a date (naive)."""
    day, month, year = date_str.split(".")
    if len(year) == 2:
        year = ("20" + year) if int(year) < 70 else ("19" + year)
    return datetime(int(year), int(month), int(day))


def parse_time(time_str: str) -> Tuple[int, int]:
    """Parse HH:MM -> (hour, minute)."""
    hour, minute = time_str.split(":")
    return int(hour), int(minute)


def stable_uid(start_dt: datetime, end_dt: datetime, location: Optional[str], description: str) -> str:
    """Create a stable UID based on key fields to avoid duplicates."""
    hasher = hashlib.sha1()
    payload = "|".join([
        start_dt.astimezone(timezone.utc).isoformat(),
        (end_dt.astimezone(timezone.utc).isoformat() if end_dt else ""),
        (location or ""),
        description,
    ])
    hasher.update(payload.encode("utf-8"))
    return hasher.hexdigest() + "@heimbas-ics"


def build_ics(entries: List[Dict[str, Any]], output_path: str) -> None:
    """Create an ICS file from parsed entries."""
    cal = Calendar()
    cal.add('prodid', '-//Heimbas Einsatz-Vorschau zu ICS//DE')
    cal.add('version', '2.0')
    cal.add('X-WR-CALNAME', 'Dienstplan')
    cal.add('X-WR-TIMEZONE', 'Europe/Berlin')

    now_utc = datetime.now(timezone.utc)

    for e in entries:
        try:
            date_naive = parse_german_date(e["date"])  # naive date
            start_h, start_m = parse_time(e["start_time"])  # type: ignore[arg-type]

            start_dt = datetime(
                date_naive.year, date_naive.month, date_naive.day,
                start_h, start_m, tzinfo=BERLIN_TZ
            )

            if e.get("end_time"):
                end_h, end_m = parse_time(e["end_time"])  # type: ignore[arg-type]
                end_dt = datetime(
                    date_naive.year, date_naive.month, date_naive.day,
                    end_h, end_m, tzinfo=BERLIN_TZ
                )
                # Prevent inverted ranges
                if end_dt <= start_dt:
                    end_dt = start_dt + timedelta(minutes=30)
            else:
                # Default duration 60 minutes if no end time
                end_dt = start_dt + timedelta(minutes=60)

            description = e.get("description", "").strip()
            address = e.get("address")
            # Title: first line or first sentence of description
            title = description.split("\n")[0]
            title = re.split(r"[\.!?]", title)[0].strip() or "Einsatz"

            vevent = Event()
            vevent.add('uid', stable_uid(start_dt, end_dt, address, description))
            vevent.add('dtstamp', now_utc)
            vevent.add('dtstart', start_dt)
            vevent.add('dtend', end_dt)
            vevent.add('summary', title)
            if address:
                vevent.add('location', address)
            if description:
                vevent.add('description', description)

            cal.add_component(vevent)
        except Exception as ex:
            debug(f"Überspringe Eintrag wegen Fehler: {ex}")
            continue

    with open(output_path, 'wb') as f:
        f.write(cal.to_ical())


def main() -> None:
    base_url = "https://homecare.hbweb.myneva.cloud/apps/cg_homecare_1017"
    username, password = get_env_credentials()

    try:
        html = login_and_get_einsatz_vorschau_html(base_url, username, password)
        entries = parse_table_entries(html)
        build_ics(entries, "dienstplan.ics")
    except RuntimeError as e:
        print(f"Fehler: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()


