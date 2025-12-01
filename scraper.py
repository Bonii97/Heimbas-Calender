import os
import sys
import re
import argparse
import hashlib
from datetime import datetime, timedelta, timezone
from typing import List, Optional, Tuple, Dict, Any

from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup
from icalendar import Calendar, Event
from zoneinfo import ZoneInfo

BERLIN_TZ = ZoneInfo("Europe/Berlin")


def debug(msg: str) -> None:
    print(f"[scraper] {msg}", file=sys.stderr)


# -----------------------------
# LOGIN + PAGE SCRAPING
# -----------------------------

def login_and_get_html(base_url: str, username: str, password: str) -> str:
    """Loggt ein und kehrt zum HTML zurück, das die Einsatz-Tabelle enthält."""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox"])
        context = browser.new_context()
        page = context.new_page()

        page.goto(base_url, timeout=60000)

        # Username
        page.fill("input[type=text]", username)
        # Password
        page.fill("input[type=password]", password)
        page.keyboard.press("Enter")
        page.wait_for_timeout(3000)

        # Fallback: Suche nach Einsatz-Vorschau / Tabelle
        html = page.content()

        browser.close()
        return html


# -----------------------------
# TABLE PARSING
# -----------------------------

def extract_date(text: str) -> Optional[str]:
    m = re.search(r"(\d{1,2}\.\d{1,2}\.(?:\d{2}|\d{4}))", text)
    return m.group(1) if m else None


def extract_time_range(text: str) -> Tuple[Optional[str], Optional[str]]:
    sanitized = re.sub(r"\b\d{1,2}\.\d{1,2}\.(?:\d{2}|\d{4})\b", " ", text)
    time_pat = r"\b\d{1,2}[:\.]\d{2}\b"

    m = re.search(rf"({time_pat})\s*[–\-]\s*({time_pat})", sanitized)
    if m:
        return m.group(1), m.group(2)

    m = re.search(rf"von\s*({time_pat})\s*bis\s*({time_pat})", sanitized, re.I)
    if m:
        return m.group(1), m.group(2)

    m = re.search(rf"({time_pat})", sanitized)
    return (m.group(1), None) if m else (None, None)


def extract_duration_minutes(text: str) -> Optional[int]:
    low = text.lower()

    m = re.search(r"(\d+)\s*(min|minute|minuten)\b", low)
    if m:
        return int(m.group(1))

    m = re.search(r"(\d{1,2})(?:[.,](\d{1,2}))?\s*(h|std|stunde|stunden)?\b", low)
    if m:
        h = int(m.group(1))
        frac = m.group(2)
        if not frac:
            return h * 60
        base = 10 if len(frac) == 1 else 100
        return h * 60 + int(round((int(frac) / base) * 60))

    return None


def parse_german_date(d: str) -> datetime:
    day, month, year = d.split(".")
    if len(year) == 2:
        year = ("20" + year) if int(year) < 70 else ("19" + year)
    return datetime(int(year), int(month), int(day))


def parse_time(t: str) -> Tuple[int, int]:
    sep = ":" if ":" in t else "."
    h, m = t.split(sep)
    return int(h), int(m)


def parse_table(html: str) -> List[Dict[str, Any]]:
    soup = BeautifulSoup(html, "lxml")
    tables = soup.find_all("table")
    if not tables:
        raise RuntimeError("Keine Tabelle gefunden.")

    chosen = None
    for table in tables:
        text = table.get_text(" ", strip=True).lower()
        if "datum" in text and "einsatz" in text:
            chosen = table
            break

    if not chosen:
        raise RuntimeError("Kein passender Einsatz-Tabellenblock erkannt.")

    rows = chosen.find_all("tr")
    entries = []

    for row in rows:
        cells = [c.get_text(" ", strip=True) for c in row.find_all(["td", "th"])]
        if len(cells) < 2:
            continue

        combined = " | ".join(cells)

        date = extract_date(combined)
        start, end = extract_time_range(combined)
        duration = extract_duration_minutes(combined)

        if not date or not start:
            continue

        entries.append({
            "date": date,
            "start_time": start,
            "end_time": end,
            "duration_minutes": duration,
            "description": combined
        })

    return entries


# -----------------------------
# ICS GENERATION
# -----------------------------

def stable_uid(start_dt, end_dt, description) -> str:
    raw = f"{start_dt}|{end_dt}|{description}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest() + "@heimbas"


def build_ics(entries: List[Dict[str, Any]], output_path: str) -> None:
    cal = Calendar()
    cal.add("prodid", "-//Heimbas ICS//DE")
    cal.add("version", "2.0")

    now_utc = datetime.now(timezone.utc)

    for e in entries:
        d = parse_german_date(e["date"])
        sh, sm = parse_time(e["start_time"])
        start_dt = datetime(d.year, d.month, d.day, sh, sm, tzinfo=BERLIN_TZ)

        if e.get("duration_minutes"):
            end_dt = start_dt + timedelta(minutes=e["duration_minutes"])
        elif e.get("end_time"):
            eh, em = parse_time(e["end_time"])
            end_dt = datetime(d.year, d.month, d.day, eh, em, tzinfo=BERLIN_TZ)
        else:
            end_dt = start_dt + timedelta(minutes=60)

        event = Event()
        summary = e["description"].split("|")[0].strip()

        event.add("uid", stable_uid(start_dt, end_dt, summary))
        event.add("dtstamp", now_utc)
        event.add("dtstart", start_dt)
        event.add("dtend", end_dt)
        event.add("summary", summary)
        event.add("description", e["description"])

        cal.add_component(event)

    with open(output_path, "wb") as f:
        f.write(cal.to_ical())


# -----------------------------
# MAIN (kein Multi-User!)
# -----------------------------

def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("--base-url", default="https://homecare.hbweb.myneva.cloud/apps/cg_homecare_1017")
    parser.add_argument("--user")
    parser.add_argument("--pass")
    parser.add_argument("--output", default="dienstplan.ics")
    return parser.parse_args()


def main():
    args = parse_args()

    username = args.user or os.environ.get("HEIMBAS_USER", "")
    password = args.password or os.environ.get("HEIMBAS_PASS", "")

    if not username or not password:
        print("Fehler: Zugangsdaten fehlen.", file=sys.stderr)
        sys.exit(2)

    html = login_and_get_html(args.base_url, username, password)
    entries = parse_table(html)
    build_ics(entries, args.output)


if __name__ == "__main__":
    main()
