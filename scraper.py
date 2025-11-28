import os
import sys
import re
import json
import argparse
import hashlib
import requests
from datetime import datetime, timedelta, timezone
from typing import List, Optional, Tuple, Dict, Any

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from bs4 import BeautifulSoup
from icalendar import Calendar, Event
from zoneinfo import ZoneInfo


BERLIN_TZ = ZoneInfo("Europe/Berlin")
# Ergaenzung (Codex): Webhook-Konfiguration fuer Google Sheets.
GSHEETS_WEBHOOK = os.environ.get("GSHEETS_WEBHOOK", "").strip()

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
                # For GWT/BBj widgets, try different fill methods
                first_element = locator.first
                try:
                    # Method 1: Standard fill
                    first_element.fill(value)
                    page.wait_for_timeout(500)  # Short wait for BBj
                    return True
                except Exception:
                    try:
                        # Method 2: Click, clear, type (für BBj-Widgets)
                        first_element.click()
                        page.wait_for_timeout(200)
                        first_element.press('Control+a')  # Select all
                        page.wait_for_timeout(100)
                        first_element.type(value, delay=50)  # Slow typing for old systems
                        page.wait_for_timeout(300)
                        return True
                    except Exception:
                        try:
                            # Method 3: Focus and direct keyboard input
                            first_element.focus()
                            page.wait_for_timeout(200)
                            page.keyboard.press('Control+a')
                            page.wait_for_timeout(100)
                            page.keyboard.type(value, delay=100)  # Very slow for compatibility
                            page.wait_for_timeout(300)
                            return True
                        except Exception:
                            try:
                                # Method 4: JavaScript value assignment (last resort)
                                page.evaluate(f"""
                                    (function() {{
                                        const el = document.querySelector('{sel}');
                                        if (el) {{
                                            el.value = '{value}';
                                            el.dispatchEvent(new Event('input', {{ bubbles: true }}));
                                            el.dispatchEvent(new Event('change', {{ bubbles: true }}));
                                            return true;
                                        }}
                                        return false;
                                    }})()
                                """)
                                page.wait_for_timeout(500)
                                return True
                            except Exception:
                                continue
        except Exception:
            continue
    return False


def try_click(page, selectors_or_text: List[str], timeout_ms: int = 1200) -> bool:
    """Try to click either css/xpath selectors or buttons/links by text."""
    for sel in selectors_or_text:
        try:
            if sel.startswith("css="):
                page.locator(sel[len("css="):]).first.click(timeout=timeout_ms)
                return True
            if sel.startswith("xpath="):
                page.locator(sel).first.click(timeout=timeout_ms)
                return True
            # Fallback: try get_by_role or get_by_text
            btn = page.get_by_role("button", name=re.compile(sel, re.I))
            if btn.count() > 0:
                btn.first.click(timeout=timeout_ms)
                return True
            link = page.get_by_role("link", name=re.compile(sel, re.I))
            if link.count() > 0:
                link.first.click(timeout=timeout_ms)
                return True
            # Try visible text anywhere
            el = page.get_by_text(re.compile(sel, re.I))
            if el.count() > 0:
                el.first.click(timeout=timeout_ms)
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
        # Launch browser with more realistic settings
        browser = p.chromium.launch(
            headless=True,
            args=[
                '--no-sandbox',
                '--disable-blink-features=AutomationControlled',
                '--disable-web-security',
                '--disable-features=VizDisplayCompositor'
            ]
        )
        context = browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            viewport={'width': 1920, 'height': 1080}
        )
        page = context.new_page()
        
        # Remove webdriver property that BBj might detect
        page.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined,
            });
        """)

        debug("Rufe Login-Seite auf…")
        page.goto(base_url, wait_until="domcontentloaded", timeout=60_000)
        initial_url = page.url  # Speichere URL für Login-Erfolg-Prüfung
        debug(f"Aktuelle URL nach erstem Load: {initial_url}")

        # Try to locate username/password fields in a robust way
        user_selectors = [
            # Standard-Selektoren
            'input[name="username"]',
            'input[name="user"]',
            'input[name="benutzer"]',
            'input[name="benutzername"]',
            'input[name="login"]',
            'input[name="email"]',
            'input[id*="user" i]',
            'input[id*="benutzer" i]',
            'input[id*="login" i]',
            'input[placeholder*="utzer" i]',
            'input[placeholder*="name" i]',
            'input[placeholder*="login" i]',
            'input[type="email"]',
            # GWT/BBj-spezifische Selektoren (Heimbas)
            'input.gwt-TextBox[type="text"]',
            'input.BBjInputE[type="text"]',
            'input.BBjControl[type="text"]',
            'input[class*="gwt-TextBox"][type="text"]',
            'input[class*="BBjInputE"][type="text"]',
            # Heimbas-spezifische Selektoren
            'input[maxlength="20"][type="text"]',  # Aus deinem Beispiel
            'input[autocomplete="off"][type="text"]',
            'input[type="text"]',  # Fallback: alle Text-Inputs
            'table input[type="text"]',
        ]
        pass_selectors = [
            # Standard-Selektoren
            'input[name="password"]',
            'input[name="pass"]',
            'input[name="passwort"]',
            'input[name="kennwort"]',
            'input[id*="pass" i]',
            'input[id*="wort" i]',
            'input[placeholder*="ass" i]',
            'input[placeholder*="wort" i]',
            # GWT/BBj-spezifische Selektoren (Heimbas)
            'input.gwt-TextBox[type="password"]',
            'input.BBjInputE[type="password"]',
            'input.BBjControl[type="password"]',
            'input[class*="gwt-TextBox"][type="password"]',
            'input[class*="BBjInputE"][type="password"]',
            'input[type="password"]',  # Fallback: alle Password-Inputs
            'table input[type="password"]',
        ]

        # Fill username/password if present on this page
        login_attempts = 0
        max_login_attempts = 3
        
        while login_attempts < max_login_attempts:
            try:
                # Für alte BBj-Systeme länger warten
                page.wait_for_timeout(5000)  # Länger warten für alte Systeme
                debug(f"Suche Login-Felder (Versuch {login_attempts + 1})…")
                
                # Zusätzlich auf BBj-spezifische Initialisierung warten
                try:
                    page.wait_for_function("document.readyState === 'complete'", timeout=10000)
                    page.wait_for_function("window.BBj || window.BBjLoaded || document.querySelector('.BBjControl')", timeout=5000)
                    debug("BBj-Framework erkannt und geladen")
                except Exception:
                    debug("BBj-Framework-Check übersprungen")
                
                # Erweiterte Selektoren für verschiedene Login-Systeme
                extended_user_selectors = user_selectors + [
                    'input[id="username"]',
                    'input[id="benutzer"]',
                    'input[id="login"]',
                    'input[id="email"]',
                    'input[class*="user"]',
                    'input[class*="benutzer"]',
                    'input[class*="login"]',
                    # Fallback: alle Text-Inputs
                    'form input[type="text"]:first-of-type',
                    'td:contains("Benutzer") + td input',
                    'td:contains("User") + td input',
                ]
                
                extended_pass_selectors = pass_selectors + [
                    'input[id="password"]',
                    'input[id="passwort"]',
                    'input[class*="pass"]',
                    'input[class*="wort"]',
                    # Fallback: alle Password-Inputs
                    'form input[type="password"]:first-of-type',
                    'td:contains("Passwort") + td input',
                    'td:contains("Password") + td input',
                ]
                
                filled_user = try_fill(page, extended_user_selectors, username)
                filled_pass = try_fill(page, extended_pass_selectors, password)

                if filled_user and filled_pass:
                    debug("Login-Felder gefüllt, klicke auf Anmelden…")
                    clicked = try_click(page, [
                        "Anmelden", "Login", "Einloggen", "Anmeldung", "Sign in", "Submit",
                        'css=button[type="submit"]',
                        'css=input[type="submit"]',
                        'css=button:has-text("Anmelden")',
                        'css=button:has-text("Login")',
                        'css=button:has-text("Submit")',
                        'css=button[class*="login"]',
                        'css=button[class*="submit"]',
                        # GWT/BBj-spezifische Button-Selektoren (Heimbas)
                        'css=button.gwt-Button',
                        'css=input.gwt-Button',
                        'css=button.BBjButton',
                        'css=input.BBjButton',
                        'css=button[class*="gwt-Button"]',
                        'css=input[class*="gwt-Button"]',
                        'css=button[class*="BBjButton"]',
                        'css=input[class*="BBjButton"]',
                        # Heimbas-spezifische Button-Selektoren
                        'css=table button',
                        'css=table input[type="button"]',
                        'css=td:contains("Anmelden") button',
                        'css=td:contains("Anmelden") input',
                    ])
                    if not clicked:
                        debug("Kein Login-Button gefunden, versuche Enter in Passwort-Feld…")
                        try:
                            page.locator(extended_pass_selectors[0]).press("Enter")
                        except Exception:
                            pass
                    
                    # Smart polling: Wait for login success indicators
                    debug("Login-Button geklickt - prüfe auf Erfolg…")
                    login_success = False
                    for attempt in range(10):  # Max 10s (10 * 1s)
                        page.wait_for_timeout(1000)
                        current_url = page.url
                        current_html = page.content()
                        
                        # Login-Erfolg-Indikatoren
                        success_indicators = [
                            current_url != initial_url,  # URL hat sich geändert
                            "anmeldung" not in current_html.lower(),  # Kein Login-Screen mehr
                            "benutzer" not in current_html.lower() or "einsatz" in current_html.lower(),  # Entweder kein Login-Feld oder Einsatz-Inhalte
                            "menu" in current_html.lower() or "navigation" in current_html.lower(),  # Menü erschienen
                            contains_einsatz_table(current_html)  # Direkt zur Tabelle
                        ]
                        
                        if any(success_indicators):
                            debug(f"Login erfolgreich nach {attempt + 1}s - URL: {current_url}")
                            login_success = True
                            break
                        
                        debug(f"Login-Check {attempt + 1}/10 - warte weiter…")
                    
                    if not login_success:
                        debug("Login-Erfolg nicht erkannt - setze trotzdem fort")
                    
                    debug(f"URL nach Login: {page.url}")
                    
                    # Sofort nach Login: Prüfe auf vorhandene Tabellen
                    debug("Prüfe Seite direkt nach Login auf Einsatz-Tabellen…")
                    if contains_einsatz_table(page.content()):
                        debug("Einsatz-Tabelle bereits auf Login-Zielseite gefunden!")
                        return page.content()  # Direkt zurückgeben, keine Navigation nötig
                    
                    break  # Login successful
                else:
                    # Check if we're already logged in or if page changed
                    current_html = page.content()
                    debug("Prüfe aktuelle Seite auf Login-Status und Tabellen…")
                    
                    # Prüfe zuerst auf Einsatz-Tabellen
                    if contains_einsatz_table(current_html):
                        debug("Einsatz-Tabelle bereits ohne Login gefunden!")
                        return current_html
                    
                    # Dann prüfe auf Login-Status
                    if any(keyword in current_html.lower() for keyword in [
                        "einsatz", "vorschau", "dienstplan", "schichtplan", "dashboard", "home", "nachrichten"
                    ]):
                        debug("Scheint bereits eingeloggt zu sein oder Login nicht erforderlich")
                        break
                    
                    debug("Kein Login-Formular gefunden, warte auf dynamische Inhalte…")
                    login_attempts += 1
                    if login_attempts < max_login_attempts:
                        page.wait_for_timeout(2000)

            except PlaywrightTimeoutError:
                debug("Timeout beim Login-Versuch")
                login_attempts += 1

        # Vor Navigation: Prüfe nochmals den aktuellen Seiteninhalt
        debug("Prüfe Seiteninhalt vor Navigation…")
        current_html = page.content()
        if contains_einsatz_table(current_html):
            debug("Einsatz-Tabelle bereits vor Navigation gefunden!")
            return current_html
        
        # Schnelle, deterministische Navigation zur Einsatz-Vorschau
        debug("Versuche zur Seite 'Einsatz-Vorschau' zu wechseln…")
        navigate_to_einsatz_vorschau(page)

        # Stelle vor dem Auslesen sicher: Zeitraum = 6 Monate (auch in Frames)
        # Exakt in dem Dokument/Frame setzen, wo die Tabelle liegt
        target_doc = find_frame_with_einsatz_table(page) or page
        try:
            set_time_range_to_six_months(target_doc)
        except Exception:
            pass

        # As a final fallback, check the current page AND frames thoroughly for any tables
        debug("Prüfe aktuelle Seite auf alle vorhandenen Tabellen…")
        current_html = page.content()
        soup = BeautifulSoup(current_html, "lxml")
        all_tables = soup.find_all("table")
        
        debug(f"Gefundene Tabellen: {len(all_tables)}")
        for i, table in enumerate(all_tables):
            # Get table text preview
            table_text = table.get_text(" ", strip=True)[:200]
            debug(f"Tabelle {i+1}: {table_text}...")
            
            # Check if this table might contain schedule data
            if any(keyword in table_text.lower() for keyword in [
                "datum", "uhrzeit", "zeit", "von", "bis", "einsatz", "adresse", 
                "kunde", "patient", "termin", "arbeitszeit"
            ]):
                debug(f"Tabelle {i+1} könnte Einsatz-Daten enthalten!")

        # Zusätzlich: In Frames nach Tabellen suchen
        for frm in page.frames:
            if frm == page.main_frame:
                continue
            try:
                frm_html = frm.content()
                frm_soup = BeautifulSoup(frm_html, "lxml")
                frm_tables = frm_soup.find_all("table")
                debug(f"Frame {getattr(frm, 'url', lambda: 'n/a')() if hasattr(frm, 'url') else 'n/a'}: {len(frm_tables)} Tabellen")
                if frm_tables:
                    # Nutze den Frame-HTML, wenn Tabelle plausibel aussieht
                    if contains_einsatz_table(frm_html):
                        debug("Plausible Einsatz-Tabelle im Frame gefunden – verwende Frame-HTML")
                        current_html = frm_html
                        final_html = frm_html
                        soup = frm_soup
                        all_tables = frm_tables
                        break
            except Exception:
                continue

        # Final fallback: Check current page content regardless of navigation success
        debug("Finale Prüfung der aktuellen Seite...")
        final_html = page.content()
        
        # Analyze all tables found on final page
        soup = BeautifulSoup(final_html, "lxml")
        all_tables = soup.find_all("table")
        debug(f"Finale Analyse: {len(all_tables)} Tabellen auf der Seite gefunden")
        
        if all_tables:
            for i, table in enumerate(all_tables):
                table_text = table.get_text(" ", strip=True)[:300]  # Erweitert für mehr Context
                debug(f"Tabelle {i+1} Inhalt: {table_text}...")
                
                # Erweiterte Keyword-Suche für Einsatz-Daten
                if any(keyword in table_text.lower() for keyword in [
                    "datum", "uhrzeit", "zeit", "von", "bis", "einsatz", "adresse", 
                    "kunde", "patient", "termin", "arbeitszeit", "dienst", "schicht",
                    "montag", "dienstag", "mittwoch", "donnerstag", "freitag", "samstag", "sonntag",
                    "januar", "februar", "märz", "april", "mai", "juni", "juli", "august", "september", "oktober", "november", "dezember"
                ]):
                    debug(f"Tabelle {i+1} enthält potentielle Einsatz-Daten - verwende sie!")
                    # Force return this table even if our detection failed
                    return final_html

        html = page.content()
        debug(f"Finale URL: {page.url}")
        context.close()
        browser.close()

    # Basic validation: ensure table likely exists
    if not contains_einsatz_table(html):
        with open("lastpage.html", "w", encoding="utf-8") as f:
            f.write(html)
        # Additional debugging: save page title and URL info
        soup = BeautifulSoup(html, "lxml")
        title = soup.find("title")
        title_text = title.get_text(strip=True) if title else "Kein Titel"
        debug(f"Seitentitel: {title_text}")
        debug(f"HTML-Länge: {len(html)} Zeichen")
        debug(f"Tabellen gefunden: {len(soup.find_all('table'))}")
        
        raise RuntimeError(
            f"Konnte keine Einsatz-Tabelle finden. Seitentitel: '{title_text}'. Die zuletzt geladene Seite wurde als 'lastpage.html' gespeichert."
        )
    return html


def contains_einsatz_table(html: str) -> bool:
    """Heuristically determine whether HTML contains the desired table."""
    soup = BeautifulSoup(html, "lxml")
    for table in soup.find_all("table"):
        header_text = " ".join(th.get_text(strip=True) for th in table.find_all(["th", "td"]))
        header_text_lower = header_text.lower()
        
        # Keywords basierend auf Screenshot der finalen Tabelle
        einsatz_keywords = [
            "datum", "einsatz", "training", "uhrzeit", "uhrzeitvon", "von", "bis", "dauer",
            "beschreibung", "adresse",
            # Spezifische Inhalte aus Screenshot
            "kohle", "martin", "feger", "nicole", "berger", "ricarda",
            "hochfellstraße", "hochriesstraße", "sommerlandstraße",
            "apvoll", "kpstd", "pbstd", "hhstd", "anfahrtspauschale",
            # Wochentage (kurz)
            "mo", "di", "mi", "do", "fr", "sa", "so"
        ]
        
        # Prüfe ob mindestens 2 Keywords gefunden werden
        found_count = sum(1 for kw in einsatz_keywords if kw in header_text_lower)
        if found_count >= 2:
            return True
            
        # Spezielle Kombinationen für Einsatz-Vorschau
        if "datum" in header_text_lower and ("einsatz" in header_text_lower or "training" in header_text_lower):
            return True
        if "von" in header_text_lower and "bis" in header_text_lower and "dauer" in header_text_lower:
            return True
    return False


def find_frame_with_einsatz_table(page) -> Optional[Any]:
    """Finde das Dokument/Frame, das die Einsatz-Tabelle enthält.

    Gibt das passende Frame/Page-Objekt zurück oder None.
    """
    try:
        # Hauptdokument zuerst prüfen
        if contains_einsatz_table(page.content()):
            return page
    except Exception:
        pass
    # Danach Frames prüfen
    try:
        for frm in page.frames:
            try:
                if contains_einsatz_table(frm.content()):
                    return frm
            except Exception:
                continue
    except Exception:
        pass
    return None


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

    # Versuche die Spaltenüberschrift für "Dauer" zu finden, um die korrekte Zelle auszulesen
    duration_col_idx: Optional[int] = None
    header_cells = [c.get_text("\n", strip=True) for c in rows[0].find_all(["td", "th"])] if rows else []
    for idx, text in enumerate(header_cells):
        if re.search(r"\bdauer\b", text, re.I):
            duration_col_idx = idx
            break
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

        # Dauer aus entsprechender Spalte oder gesamtem Zeilentext extrahieren (Minuten)
        duration_minutes: Optional[int] = None
        if duration_col_idx is not None and len(cells) > duration_col_idx:
            duration_minutes = extract_duration_minutes(cells[duration_col_idx])
        if duration_minutes is None:
            duration_minutes = extract_duration_minutes(combined)

        if not date_str or not start_str:
            # Not enough information to build an event; skip header or invalid rows
            continue

        entries.append({
            "date": date_str,
            "start_time": start_str,
            "end_time": end_str,
            "description": description,
            "address": address,
            "duration_minutes": duration_minutes,
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
    """Extract time range like '08:00 - 09:30' or 'von 8:00 bis 9:30'.

    Vor der Zeitsuche werden Datumsteile (dd.mm.yyyy|yy) entfernt, um Fehl-Treffer
    wie '13.08' aus '13.08.25' zu vermeiden.
    """
    # Datumsteile entfernen
    sanitized = re.sub(r"\b\d{1,2}\.\d{1,2}\.(?:\d{2}|\d{4})\b", " ", text)
    # Zeitformat HH:MM oder HH.MM
    time_pat = r"\b\d{1,2}[:\.]\d{2}\b"
    # Pattern mit Gedankenstrich
    m = re.search(rf"({time_pat})\s*[–\-]\s*({time_pat})", sanitized)
    if m:
        return m.group(1), m.group(2)
    # Pattern mit 'von ... bis ...'
    m = re.search(rf"von\s*({time_pat})\s*bis\s*({time_pat})", sanitized, re.I)
    if m:
        return m.group(1), m.group(2)
    # Einzelne Zeit (Fallback)
    m = re.search(rf"({time_pat})", sanitized)
    if m:
        return m.group(1), None
    return None, None


def extract_duration_minutes(text: str) -> Optional[int]:
    """Extract duration in minutes from text.

    Unterstützte Formate:
    - "2,0" oder "2.0" (Stunden)
    - "2 Std." / "2 Stunden"
    - "90 Min" / "90 Minuten"
    - gemischte Formate in einer Zelle
    """
    t = text.strip().lower()
    # Minuten-Angaben
    m = re.search(r"(\d{1,3})\s*(min|minute|minuten)\b", t)
    if m:
        return int(m.group(1))

    # Stunden-Angaben wie 2,0 / 2,00 / 2.5 / 2 Std.
    m = re.search(r"(\d{1,2})([\.,](\d{1,2}))?\s*(h|std|stunde|stunden)?\b", t)
    if m:
        hours = int(m.group(1))
        frac_str = m.group(3)
        if frac_str is None:
            return hours * 60
        # 1 oder 2 Dezimalstellen → in Minuten umrechnen
        base = 10 if len(frac_str) == 1 else 100
        minutes = hours * 60 + int(round((int(frac_str) / base) * 60))
        return minutes

    return None


def set_time_range_to_six_months(page) -> None:
    """Wählt in der Heimbas-Oberfläche den Zeitraum '6 Monate' aus.

    Vorgehen:
    1) Suche das Dropdown/Select in der Nähe von 'Zeitraum'
    2) Versuche per sichtbarem Text '6 Monate' auszuwählen
    3) Führe intelligentes Polling durch, ob sich Tabelle erweitert/Content ändert
    """
    debug("Versuche Zeitraum '6 Monate' zu setzen…")

    # Wenn bereits aktiv, überspringen
    try:
        current_label = page.evaluate("""
            (() => {
                const el = document.querySelector('div.basis-button-face');
                return el ? (el.innerText||'').trim() : '';
            })()
        """)
        if isinstance(current_label, str) and re.search(r"6\s*Monat", current_label, re.I):
            debug("Zeitraum bereits auf '6 Monate' gesetzt – übersprungen")
            return
    except Exception:
        pass

    # Vorheriger Content-Snapshot
    before_html = page.content()

    # Versuche per Rollen-/Textauswahl (Playwright heuristics)
    try:
        # Mögliche Ansätze: select-Element, Button mit Dropdown, Text 'Zeitraum'
        # 1) Direktes Select mit sichtbarem Namen
        select_locators = [
            page.get_by_label("Zeitraum", exact=False),
            page.locator("select"),
        ]
        made_selection = False
        for sel in select_locators:
            if sel.count() > 0:
                try:
                    # versuche die Option '6 Monate' zu wählen
                    sel.first.select_option(label="6 Monate")
                    made_selection = True
                    break
                except Exception:
                    # manche Oberflächen sind keine echten <select> Elemente
                    pass

        if not made_selection:
            # 2) Klick-Pfade (Dropdown öffnen → '6 Monate' klicken)
            # Besonderheit Heimbas: zuerst auf aktuellen Wert (z. B. '7 Tage') klicken
            clicked_7 = try_click(page, [
                'css=div.basis-button-face:has-text("7 Tage")',
                '7 Tage',
                'css=button:has-text("7 Tage")',
                'css=div:has-text("7 Tage")',
                'css=span:has-text("7 Tage")',
            ])
            # Wenn '7 Tage' nicht geklickt werden konnte, Dropdown direkt öffnen
            if not clicked_7:
                try_click(page, [
                    'css=button:has-text("Zeitraum")',
                    'css=div:has-text("Zeitraum")',
                    'css=[role="button"]:has-text("Zeitraum")',
                ], timeout_ms=800)

            # Jetzt gezielt '6 Monate' mit kurzen Timeouts versuchen
            try_click(page, [
                'css=button:has-text("Zeitraum")',
                'css=div:has-text("Zeitraum")',
            ], timeout_ms=800)
            selected_6 = try_click(page, [
                'css=div.HMBListBoxItem.HMBNavButton.mynevaListBoxItemBorderLeft.mynevaListBoxItemBorderRight:has-text("6 Monate")',
                'css=li:has-text("6 Monate")',
                'css=div[role="option"]:has-text("6 Monate")',
                'css=button:has-text("6 Monate")',
                '6 Monate',
            ], timeout_ms=800)
            if not selected_6:
                debug("Option '6 Monate' nicht gefunden – Abbruch der Auswahlsequenz")
                return

        # Smart Polling: bis zu 8 Sekunden auf Änderungen warten
        # Max 4 Sekunden Polling – keine langen Hänger
        for attempt in range(4):
            page.wait_for_timeout(1000)
            after_html = page.content()
            if len(after_html) != len(before_html):
                debug(f"Zeitraum umgestellt (Änderung nach {attempt + 1}s erkannt)")
                break
        else:
            debug("Keine erkennbare Änderung nach Zeitraum-Umstellung – fahre fort")
    except Exception as ex:
        debug(f"Fehler beim Setzen des Zeitraums: {ex}")


def navigate_to_einsatz_vorschau(page) -> None:
    """Schnell und deterministisch zum Menüpunkt 'Einsatz-Vorschau' navigieren.

    Strategie:
    1) Klick per starken CSS-Selektoren mit kurzen Timeouts
    2) Falls nötig: Fallback per JS (querySelectorAll + exact match)
    3) Kurzes Polling (max. 4s) auf Tabellen/Schlüsselbegriffe
    """
    # 1) Direkte Klicks mit kurzen Timeouts
    if try_click(page, [
        'css=td:has-text("Einsatz-Vorschau")',
        'css=.HMBListBoxContent:has-text("Einsatz-Vorschau")',
        'css=div:has-text("Einsatz-Vorschau")',
        'css=span:has-text("Einsatz-Vorschau")',
        'css=a:has-text("Einsatz-Vorschau")',
        'Einsatz-Vorschau'
    ], timeout_ms=800):
        pass
    else:
        # 2) Fallback per JS
        try:
            result = page.evaluate("""
                (function() {
                    const els = document.querySelectorAll('td,div,span,a');
                    for (const el of els) {
                        const t = (el.innerText||'').trim();
                        if (t === 'Einsatz-Vorschau') { el.click(); return true; }
                    }
                    return false;
                })()
            """)
            if not result:
                debug("JS-Navigation auf 'Einsatz-Vorschau' nicht erfolgreich")
        except Exception as ex:
            debug(f"JS-Navigation Fehler: {ex}")

    # 3) kurzes Polling auf Navigationserfolg (max. 4s)
    for i in range(4):
        page.wait_for_timeout(1000)
        html = page.content()
        if contains_einsatz_table(html) or re.search(r"\b(einsatz|vorschau)\b", html, re.I):
            debug(f"Navigation erkannt nach {i+1}s")
            return
    debug("Navigation auf 'Einsatz-Vorschau' nicht sicher erkannt – fahre fort")

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
    """Parse HH:MM oder HH.MM -> (hour, minute) mit Validierung."""
    sep = ":" if ":" in time_str else "."
    hour_str, minute_str = time_str.split(sep)
    hour, minute = int(hour_str), int(minute_str)
    if not (0 <= hour <= 23 and 0 <= minute <= 59):
        raise ValueError(f"Ungültige Zeit: {time_str}")
    return hour, minute


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


def slugify_name(name: str) -> str:
    """Create a filesystem/url friendly slug from the given name."""
    # Lowercase, replace spaces with underscore, allow only a-z0-9_- characters
    s = name.strip().lower().replace(" ", "_")
    s = re.sub(r"[^a-z0-9_-]", "_", s)
    # collapse multiple underscores
    s = re.sub(r"_+", "_", s).strip("_")
    return s or "user"


# Ergaenzung (Codex): Hilfsfunktionen fuer den Webhook-Export.
def make_einsatz_id_from_entry(entry: Dict[str, Any]) -> str:
    """Create a stable SHA1 over the essential plan fields."""
    date_str = str(entry.get("date", "")).strip()
    start_str = str(entry.get("start_time", "")).strip()
    end_str = str(entry.get("end_time", "")).strip()
    address = str(entry.get("address", "") or "").strip()
    title = str(entry.get("title") or "").strip()
    if not title:
        description = str(entry.get("description", "")).strip()
        first_line = description.split("\n")[0]
        title = re.split(r"[\.!?]", first_line)[0].strip() or "Einsatz"
    raw = "|".join([date_str, start_str, end_str, address, title])
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()


def send_plan_records(user_label: str, entries: List[Dict[str, Any]]) -> None:
    """Send plan entries to the configured Google Sheets webhook."""
    if not GSHEETS_WEBHOOK or not entries:
        return
    for entry in entries:
        try:
            date_str = str(entry.get("date", "")).strip()
            title_text = str(entry.get("title") or "").strip()
            if not title_text:
                description = str(entry.get("description", "")).strip()
                first_line = description.split("\n")[0]
                title_text = re.split(r"[\.!?]", first_line)[0].strip() or "Einsatz"
            entry_for_hash = dict(entry)
            entry_for_hash.setdefault("title", title_text)
            einsatz_id = make_einsatz_id_from_entry(entry_for_hash)
            try:
                german_date = parse_german_date(date_str).strftime("%d.%m.%Y")
            except Exception:
                german_date = ""

            start_str = str(entry.get("start_time", "")).strip()
            if start_str.lower() == "none":
                start_str = ""
            raw_end = entry.get("end_time")
            end_str = str(raw_end).strip() if raw_end is not None else ""
            if end_str.lower() == "none":
                end_str = ""

            if not end_str:
                duration_minutes = entry.get("duration_minutes")
                try:
                    start_dt = parse_german_date(date_str)
                    start_hour, start_minute = parse_time(start_str)
                    duration = duration_minutes if isinstance(duration_minutes, int) and duration_minutes > 0 else 60
                    computed_end = datetime(
                        start_dt.year,
                        start_dt.month,
                        start_dt.day,
                        start_hour,
                        start_minute,
                    ) + timedelta(minutes=duration)
                    end_str = f"{computed_end.hour:02d}:{computed_end.minute:02d}"
                except Exception:
                    end_str = ""
            payload = {
                "type": "plan",
                "user": user_label,
                "einsatzId": einsatz_id,
                "datum": german_date or date_str,
                "start_plan": start_str,
                "ende_plan": end_str,
                "titel": title_text,
                "adresse": str(entry.get("address", "") or "").strip(),
            }
            response = requests.post(GSHEETS_WEBHOOK, json=payload, timeout=10)
            if response.status_code >= 400:
                debug(f"Webhook-Fehler ({response.status_code}) fuer Einsatz {einsatz_id}")
        except requests.RequestException as req_err:
            debug(f"Webhook nicht erreichbar: {req_err}")
        except Exception as err:
            debug(f"Webhook payload Fehler: {err}")


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

            # Priorität: wenn eine Endzeit angegeben ist, verwende sie; sonst Dauer
            explicit_duration_min: Optional[int] = e.get("duration_minutes")  # type: ignore[assignment]

            if e.get("end_time"):
                end_h, end_m = parse_time(e["end_time"])  # type: ignore[arg-type]
                end_dt = datetime(
                    date_naive.year, date_naive.month, date_naive.day,
                    end_h, end_m, tzinfo=BERLIN_TZ
                )
                # Prevent inverted ranges
                if end_dt <= start_dt:
                    end_dt = start_dt + timedelta(minutes=30)
            elif explicit_duration_min is not None and explicit_duration_min > 0:
                end_dt = start_dt + timedelta(minutes=explicit_duration_min)
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


def fetch_entries_for_user(base_url: str, username: str, password: str) -> List[Dict[str, Any]]:
    html = login_and_get_einsatz_vorschau_html(base_url, username, password)
    return parse_table_entries(html)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Heimbas Einsatz-Vorschau zu ICS")
    parser.add_argument("--base-url", default="https://homecare.hbweb.myneva.cloud/apps/cg_homecare_1017",
                        help="Basis-URL der Heimbas-Anwendung")
    parser.add_argument("--user", dest="user", help="Benutzername (überschreibt HEIMBAS_USER)")
    parser.add_argument("--pass", dest="password", help="Passwort (überschreibt HEIMBAS_PASS)")
    parser.add_argument("--output", dest="output", default="dienstplan.ics", help="Ausgabepfad der ICS")
    parser.add_argument("--users-json-path", dest="users_json_path",
                        help="Pfad zu einer JSON-Datei mit mehreren Accounts [{name,user,pass}]")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    base_url = args.base_url

    # Multi-User-Modus: via Datei oder Umgebungsvariable USERS_JSON
    users_json_text = None
    if args.users_json_path and os.path.exists(args.users_json_path):
        with open(args.users_json_path, "r", encoding="utf-8") as f:
            users_json_text = f.read()
    elif os.environ.get("USERS_JSON"):
        users_json_text = os.environ["USERS_JSON"]

    if users_json_text:
        try:
            users_list = json.loads(users_json_text)
            if not isinstance(users_list, list):
                raise ValueError("USERS_JSON ist kein Array.")
        except Exception as ex:
            print(f"Fehler beim Lesen von USERS_JSON: {ex}", file=sys.stderr)
            sys.exit(2)

        combined_entries: List[Dict[str, Any]] = []
        any_success = False
        for idx, entry in enumerate(users_list):
            # Unterstütze verschiedene Key-Varianten: name/label, user/username, pass/password
            name_raw = str(
                (entry.get("name") or entry.get("label") or f"user{idx+1}")
            ).strip()
            name = slugify_name(name_raw)
            u = str((entry.get("user") or entry.get("username") or "")).strip()
            p = str((entry.get("pass") or entry.get("password") or "")).strip()
            if not u or not p:
                debug(f"Eintrag '{name_raw}' hat keine vollständigen Zugangsdaten – übersprungen.")
                continue
            try:
                debug(f"Lese Einsätze für '{name_raw}' (Datei-Slug: '{name}')…")
                user_entries = fetch_entries_for_user(base_url, u, p)
                # Ergaenzung (Codex): Webhook pro Benutzer direkt nach dem Abruf.
                send_plan_records(user_label=name, entries=user_entries)
                combined_entries.extend(user_entries)
                # pro User eigene Datei
                out_user_path = f"dienstplan_{name}.ics"
                build_ics(user_entries, out_user_path)
                any_success = True
            except Exception as ex:
                debug(f"Fehler für Benutzer '{name_raw}': {ex}")
                continue

        if not any_success:
            print("Fehler: Konnte für keinen Benutzer Einsätze erzeugen.", file=sys.stderr)
            sys.exit(1)

        # Kombinierte Datei
        if combined_entries:
            build_ics(combined_entries, args.output)
        return

    # Single-User-Modus
    username = args.user or os.environ.get("HEIMBAS_USER", "").strip()
    password = args.password or os.environ.get("HEIMBAS_PASS", "").strip()
    if not username or not password:
        print("Fehler: Zugangsdaten fehlen (Argumente oder Umgebungsvariablen).", file=sys.stderr)
        sys.exit(2)

    try:
        entries = fetch_entries_for_user(base_url, username, password)
        # Ergaenzung (Codex): Webhook fuer den Single-User-Ausgang.
        send_plan_records(user_label=slugify_name(username), entries=entries)
        build_ics(entries, args.output)
    except RuntimeError as e:
        print(f"Fehler: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
