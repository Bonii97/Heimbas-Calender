import os
import sys
import re
import json
import argparse
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
        
        # Try to navigate/click to 'Einsatz-Vorschau'
        debug("Versuche zur Seite 'Einsatz-Vorschau' zu wechseln…")
        
        # First, try to find existing navigation in the current page
        nav_attempts = 0
        max_nav_attempts = 2  # Reduziert von 3 auf 2
        
        while nav_attempts < max_nav_attempts:
            # Look for navigation elements in current page - basierend auf Screenshots
            nav_clicked = try_click(page, [
                # Exact text match from screenshot (linkes Menü, zweiter Eintrag)
                "Einsatz-Vorschau",
                # Priorität auf linkes Menü-Element (Screenshot zeigt Menü-Navigation)
                'css=td:has-text("Einsatz-Vorschau")',  # Table-basierte Navigation (häufig in alten Systems)
                'css=div:has-text("Einsatz-Vorschau")',
                'css=span:has-text("Einsatz-Vorschau")',
                'css=a:has-text("Einsatz-Vorschau")',
                # Menü-basierte Selektoren
                'css=.menu td:has-text("Einsatz-Vorschau")',
                'css=.navigation td:has-text("Einsatz-Vorschau")',
                'css=table td:has-text("Einsatz-Vorschau")',
                # Heimbas-specific selectors (falls framework-spezifisch)
                'css=div.HMBListBoxContent:has-text("Einsatz-Vorschau")',
                'css=div[class*="HMBListBox"]:has-text("Einsatz-Vorschau")',
                # Generic fallbacks
                "Einsatz Vorschau", "Vorschau", "Einsatz",
                # Link/button selectors
                'css=a[href*="einsatz"]',
                'css=a[href*="vorschau"]',
                'css=button[onclick*="einsatz"]',
                'css=button[onclick*="vorschau"]',
            ])
            
            if nav_clicked:
                debug("Navigation-Element gefunden und geklickt")
                try:
                    # Smart polling: Wait for navigation success
                    debug("Prüfe auf Navigation-Erfolg…")
                    nav_success = False
                    
                    for attempt in range(8):  # Max 8s (8 * 1s)
                        page.wait_for_timeout(1000)
                        current_html = page.content()
                        
                        # Navigation-Erfolg-Indikatoren
                        nav_indicators = [
                            contains_einsatz_table(current_html),  # Ziel-Tabelle gefunden
                            "einsatz-vorschau" in current_html.lower(),  # Seiten-Inhalt passt
                            "datum" in current_html.lower() and "uhrzeit" in current_html.lower(),  # Tabellen-Header
                            len(BeautifulSoup(current_html, "lxml").find_all("table")) > 0  # Mindestens eine Tabelle
                        ]
                        
                        if any(nav_indicators):
                            debug(f"Navigation erfolgreich nach {attempt + 1}s")
                            nav_success = True
                            break
                        
                        debug(f"Navigation-Check {attempt + 1}/8 - warte weiter…")
                    
                    # Detaillierte Analyse nach Navigation (erfolgreich oder nicht)
                    current_html = page.content()
                    soup = BeautifulSoup(current_html, "lxml")
                    
                    debug(f"URL nach Navigation: {page.url}")
                    
                    # Log page info for debugging
                    title = soup.find("title")
                    title_text = title.get_text(strip=True) if title else "Kein Titel"
                    debug(f"Seitentitel: {title_text}")
                    
                    # Prüfe sofort auf Tabellen
                    if contains_einsatz_table(current_html):
                        debug("Einsatz-Tabelle nach Navigation gefunden!")
                        break
                    
                    # Check if the navigation actually changed something
                    all_tables = soup.find_all("table")
                    debug(f"Tabellen nach Navigation: {len(all_tables)}")
                    
                    if len(all_tables) > 0:
                        for i, table in enumerate(all_tables):
                            table_preview = table.get_text(" ", strip=True)[:150]
                            debug(f"Tabelle {i+1} nach Navigation: {table_preview}")
                    
                except Exception as ex:
                    debug(f"Navigation-Fehler: {ex}")
                    # Continue anyway
            
            # Try JavaScript navigation for single page apps
            debug(f"Suche nach JS-Navigation (Versuch {nav_attempts + 1})…")
            try:
                # Try Heimbas-specific JavaScript navigation
                result = page.evaluate("""
                    (function() {
                        // Try exact match first
                        let found = false;
                        const elements = document.querySelectorAll('div.HMBListBoxContent, div[class*="HMBListBox"], td, span, div, a');
                        
                        for (let item of elements) {
                            const text = (item.innerText || item.textContent || '').trim();
                            if (text === 'Einsatz-Vorschau') {
                                console.log('Found exact match:', item);
                                item.click();
                                found = true;
                                break;
                            }
                        }
                        
                        if (!found) {
                            // Fallback: partial match
                            for (let item of elements) {
                                const text = (item.innerText || item.textContent || '').trim();
                                if (text.match(/(einsatz.*vorschau|vorschau.*einsatz)/i)) {
                                    console.log('Found partial match:', item);
                                    item.click();
                                    found = true;
                                    break;
                                }
                            }
                        }
                        
                        return found;
                    })()
                """)
                
                if result:
                    debug("JavaScript-Navigation durchgeführt")
                
                # Smart polling für JS-Navigation
                js_success = False
                for attempt in range(6):  # Max 6s
                    page.wait_for_timeout(1000)
                    current_html = page.content()
                    
                    if contains_einsatz_table(current_html):
                        debug(f"Einsatz-Tabelle nach JS-Navigation gefunden! ({attempt + 1}s)")
                        js_success = True
                        break
                    
                    # Weitere Erfolgs-Indikatoren für JS-Navigation
                    if ("datum" in current_html.lower() and "uhrzeit" in current_html.lower()) or \
                       len(BeautifulSoup(current_html, "lxml").find_all("table")) > 0:
                        debug(f"JS-Navigation Indikatoren erkannt nach {attempt + 1}s")
                
                if js_success:
                    break
                    
            except Exception as ex:
                debug(f"JS-Navigation fehlgeschlagen: {ex}")
            
            nav_attempts += 1
            if nav_attempts < max_nav_attempts:
                page.wait_for_timeout(3000)

        # As a final fallback, check the current page thoroughly for any tables
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
        from bs4 import BeautifulSoup
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


def slugify_name(name: str) -> str:
    """Create a filesystem/url friendly slug from the given name."""
    # Lowercase, replace spaces with underscore, allow only a-z0-9_- characters
    s = name.strip().lower().replace(" ", "_")
    s = re.sub(r"[^a-z0-9_-]", "_", s)
    # collapse multiple underscores
    s = re.sub(r"_+", "_", s).strip("_")
    return s or "user"


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
        build_ics(entries, args.output)
    except RuntimeError as e:
        print(f"Fehler: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()


