#!/usr/bin/env python3
"""
Stadtbibliothek Halle → Google Calendar Sync + Ausleih-Verlauf
===============================================================
Liest ausgeliehene Medien aller Benutzer aus dem Bibliotheksportal (katalog.halle.de),
erstellt/aktualisiert Kalendereinträge in Google Calendar und führt eine Verlaufs-
Übersicht (verlauf.xml / verlauf.html) mit Covern und Ausleihdaten.

Medien mit gleichem Rückgabedatum werden in einem gemeinsamen Kalendereintrag
zusammengefasst.

Setup-Anleitung: siehe README-Abschnitt am Ende dieser Datei.
"""

import os
import re
import time
import hashlib
import datetime
import xml.etree.ElementTree as ET
from xml.dom import minidom
from pathlib import Path
import requests
from bs4 import BeautifulSoup

# Google API
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# ─────────────────────────────────────────────
# KONFIGURATION – hier anpassen
# ─────────────────────────────────────────────

BENUTZER = [
    {
        "name":    "Laura",
        "ausweis": os.environ.get("BIBL_AUSWEIS_LAURA", ""),
        "passwort": os.environ.get("BIBL_PASSWORT_LAURA", ""),
    },
    {
        "name":    "Benny",
        "ausweis": os.environ.get("BIBL_AUSWEIS_BENNY", ""),
        "passwort": os.environ.get("BIBL_PASSWORT_BENNY", ""),
    },
]

GOOGLE_CREDENTIALS_FILE = "credentials.json"
GOOGLE_TOKEN_FILE       = "token.json"

# ID des Ziel-Kalenders (oder "primary" für Hauptkalender)
KALENDER_ID = "21b41bda46bed0a74602bc0234ff02aea277e70fec21548420e4526982a02f07@group.calendar.google.com"

# Wie viele Tage vor Ablauf soll der Erinnerungsalarm erscheinen?
ERINNERUNG_TAGE_VORHER = 3

# Verlauf-Dateien
VERLAUF_XML  = "verlauf.xml"
VERLAUF_HTML = "verlauf.html"
COVERS_DIR   = "covers"

# Optionaler Passwortschutz für verlauf.html (leer = kein Schutz)
SEITEN_PASSWORT = os.environ.get("HTML_PASSWORT", "")

# ─────────────────────────────────────────────
# KONSTANTEN
# ─────────────────────────────────────────────

LOGIN_URL = "https://katalog.halle.de/Mein-Konto"
SCOPES    = ["https://www.googleapis.com/auth/calendar"]


# ─────────────────────────────────────────────
# SCHRITT 1: Bibliothekskonto scrapen
# ─────────────────────────────────────────────

def hole_ausgeliehene_medien(name, ausweis, passwort):
    """
    Loggt sich ins Bibliotheksportal ein und liest alle ausgeliehenen Medien.
    Rückgabe: Liste von Dicts mit keys: titel, frist, medium_id, name,
              verfasser, mediengruppe, isbn
    """
    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )
    })

    print(f"📡 Verbinde mit Bibliotheksportal ({name})...")
    response = session.get(LOGIN_URL)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")

    def get_hidden(field_name):
        tag = soup.find("input", {"name": field_name})
        return tag["value"] if tag else ""

    post_data = {
        "__EVENTTARGET":        "",
        "__EVENTARGUMENT":      "",
        "__VIEWSTATE":          get_hidden("__VIEWSTATE"),
        "__VIEWSTATEGENERATOR": get_hidden("__VIEWSTATEGENERATOR"),
        "__EVENTVALIDATION":    get_hidden("__EVENTVALIDATION"),
        "__VIEWSTATEENCRYPTED": get_hidden("__VIEWSTATEENCRYPTED"),
        "dnn$ctr375$Login$Login_COP$txtUsername": ausweis,
        "dnn$ctr375$Login$Login_COP$txtPassword": passwort,
        "dnn$ctr375$Login$Login_COP$cmdLogin":    "Anmelden",
    }

    print(f"🔐 Melde an ({name})...")
    response = session.post(LOGIN_URL, data=post_data)
    response.raise_for_status()

    if "abmelden" not in response.text.lower():
        raise RuntimeError(
            f"❌ Login fehlgeschlagen für {name}. Bitte Ausweisnummer und Passwort prüfen.\n"
            "   Hinweis: run_debug_login() aufrufen für Diagnose."
        )
    print(f"✅ Login erfolgreich ({name}).")

    return parse_ausleih_seite(response.text, name)


def parse_ausleih_seite(html, name):
    """Parst die HTML-Seite und extrahiert ausgeliehene Medien."""
    soup = BeautifulSoup(html, "html.parser")
    medien = []

    tabelle = soup.find("table", {"id": lambda x: x and "grdViewLoans" in x})

    if not tabelle:
        print(f"ℹ️  Keine ausgeliehenen Medien für {name}.")
        return medien

    for zeile in tabelle.find_all("tr")[1:]:  # erste Zeile = Header überspringen
        zellen = zeile.find_all("td")
        if len(zellen) < 4:
            continue

        titel_link = zeile.find("a", href=lambda h: h and "Mediensuche" in h)
        if not titel_link:
            continue
        titel = titel_link.get_text(strip=True)

        href = titel_link.get("href", "")
        id_match = re.search(r"id=(\d+)", href)
        medium_id = id_match.group(1) if id_match else titel

        # Nehme das späteste Datum (= Rückgabefrist, nicht Ausleihdatum)
        datum = None
        for zelle in zellen:
            datum_match = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", zelle.get_text(strip=True))
            if datum_match:
                tag, monat, jahr = datum_match.groups()
                kandidat = datetime.date(int(jahr), int(monat), int(tag))
                if datum is None or kandidat > datum:
                    datum = kandidat

        # Verfasser
        verfasser = ""
        for zelle in zellen:
            label = zelle.find("span", class_=lambda c: c and "oclc-module-label" in c)
            if label and "Verfasser" in label.get_text():
                value = label.find_next_sibling("span")
                if value:
                    verfasser = value.get_text(strip=True)
                break

        # Mediengruppe
        mediengruppe = ""
        for zelle in zellen:
            label = zelle.find("span", class_=lambda c: c and "oclc-module-label" in c)
            if label and "Mediengruppe" in label.get_text():
                value = label.find_next_sibling("span")
                if value:
                    mediengruppe = value.get_text(strip=True)
                break

        # ISBN-13 aus data-devsources des Cover-Images (nur 978/979-Nummern sind Buch-ISBNs)
        isbn = ""
        cover_img = zeile.find("img", class_=lambda c: c and "coverSmall" in c)
        if cover_img:
            for attr in ("data-devsources", "data-sources"):
                match = re.search(r"GetMvbCover\|a\|0\$(\d+)", cover_img.get(attr, ""))
                if match:
                    candidate = match.group(1)
                    if candidate.startswith(("978", "979")):
                        isbn = candidate
                    break

        if titel and datum:
            medien.append({
                "titel":        titel,
                "frist":        datum,
                "medium_id":    medium_id,
                "name":         name,
                "verfasser":    verfasser,
                "mediengruppe": mediengruppe,
                "isbn":         isbn,
            })
            print(f"   📖 {titel} → Frist: {datum.strftime('%d.%m.%Y')}")

    return medien


# ─────────────────────────────────────────────
# SCHRITT 2: Google Calendar API
# ─────────────────────────────────────────────

def google_calendar_service():
    """Gibt einen authentifizierten Google Calendar Service zurück."""
    creds = None

    if os.path.exists(GOOGLE_TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                GOOGLE_CREDENTIALS_FILE, SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open(GOOGLE_TOKEN_FILE, "w") as token:
            token.write(creds.to_json())

    return build("calendar", "v3", credentials=creds)


def hole_bestehende_biblio_events(service):
    """
    Holt alle bestehenden Bibliotheks-Kalendereinträge.
    Rückgabe: (events_by_datum, alte_events)
      - events_by_datum: Dict {datum_str: event_dict}  (neues Format)
      - alte_events: Liste von Events im alten Format (pro Medium) → werden gelöscht
    """
    events_by_datum = {}
    alte_events = []
    page_token = None

    while True:
        result = service.events().list(
            calendarId=KALENDER_ID,
            privateExtendedProperty="bibliothek_sync=true",
            singleEvents=True,
            maxResults=250,
            pageToken=page_token,
        ).execute()

        for event in result.get("items", []):
            props = event.get("extendedProperties", {}).get("private", {})
            datum = props.get("bibliothek_datum")
            if datum:
                events_by_datum[datum] = event
            else:
                # Altes Format (ein Event pro Medium) → zum Migrieren vormerken
                alte_events.append(event)

        page_token = result.get("nextPageToken")
        if not page_token:
            break

    return events_by_datum, alte_events


def baue_event_body(datum_str, medien_gruppe):
    """Erstellt den Event-Body für eine Gruppe von Medien am gleichen Tag."""
    anzahl = len(medien_gruppe)
    summary = "📚 1 Medium abgeben" if anzahl == 1 else f"📚 {anzahl} Medien abgeben"

    beschreibung = "\n".join(
        f"- {m['titel']} ({m['name']})"
        for m in sorted(medien_gruppe, key=lambda m: (m["name"], m["titel"]))
    )

    erinnerung_minuten = ERINNERUNG_TAGE_VORHER * 24 * 60
    medium_ids_str = ",".join(sorted(m["medium_id"] for m in medien_gruppe))

    return {
        "summary":     summary,
        "description": beschreibung,
        "start":       {"date": datum_str},
        "end":         {"date": datum_str},
        "reminders": {
            "useDefault": False,
            "overrides": [
                {"method": "popup", "minutes": erinnerung_minuten},
                {"method": "email", "minutes": erinnerung_minuten},
            ],
        },
        "extendedProperties": {
            "private": {
                "bibliothek_datum": datum_str,
                "bibliothek_sync":  "true",
                "medium_ids":       medium_ids_str,
            }
        },
        "colorId": "9",  # Blaubeere
    }


def sync_events(service, medien_nach_datum, bestehende_events):
    """Erstellt/aktualisiert Events und löscht veraltete."""
    for datum_str, medien_gruppe in sorted(medien_nach_datum.items()):
        neues_body = baue_event_body(datum_str, medien_gruppe)
        neue_ids = neues_body["extendedProperties"]["private"]["medium_ids"]

        if datum_str in bestehende_events:
            altes_event = bestehende_events[datum_str]
            alte_ids = altes_event.get("extendedProperties", {}).get("private", {}).get("medium_ids", "")

            if alte_ids == neue_ids:
                print(f"   ⏭️  Unverändert: {datum_str} ({len(medien_gruppe)} Medium/Medien)")
            else:
                service.events().update(
                    calendarId=KALENDER_ID,
                    eventId=altes_event["id"],
                    body=neues_body,
                ).execute()
                print(f"   🔄 Aktualisiert: {datum_str} ({len(medien_gruppe)} Medium/Medien)")
        else:
            service.events().insert(
                calendarId=KALENDER_ID,
                body=neues_body,
            ).execute()
            titel_liste = ", ".join(f"{m['titel']} ({m['name']})" for m in medien_gruppe)
            print(f"   ✅ Neu: {datum_str} → {titel_liste}")

    for datum_str, event in bestehende_events.items():
        if datum_str not in medien_nach_datum:
            service.events().delete(
                calendarId=KALENDER_ID,
                eventId=event["id"],
            ).execute()
            print(f"   🗑️  Gelöscht (alle Medien zurückgegeben): {datum_str}")


# ─────────────────────────────────────────────
# SCHRITT 3: Verlauf-Tracking
# ─────────────────────────────────────────────

def _xml_text(element, tag):
    child = element.find(tag)
    return (child.text or "") if child is not None else ""


def hole_cover_url(isbn):
    """Sucht Cover bei Open Library (primär) oder Google Books (Fallback)."""
    if not isbn:
        return ""

    # Primär: Open Library (höhere Auflösung)
    ol_url = f"https://covers.openlibrary.org/b/isbn/{isbn}-L.jpg"
    try:
        resp = requests.head(ol_url, params={"default": "false"}, timeout=8, allow_redirects=True)
        if resp.status_code == 200:
            return ol_url
    except Exception as exc:
        print(f"      ⚠️  Open Library fehlgeschlagen für ISBN {isbn}: {exc}")

    # Fallback: Google Books
    try:
        resp = requests.get(
            "https://www.googleapis.com/books/v1/volumes",
            params={"q": f"isbn:{isbn}", "maxResults": 1},
            timeout=10,
        )
        resp.raise_for_status()
        items = resp.json().get("items", [])
        if items:
            links = items[0].get("volumeInfo", {}).get("imageLinks", {})
            url = links.get("thumbnail") or links.get("smallThumbnail") or ""
            if url:
                return url.replace("http://", "https://")
    except Exception as exc:
        print(f"      ⚠️  Google Books fehlgeschlagen für ISBN {isbn}: {exc}")

    print(f"      ⚠️  Kein Cover gefunden für ISBN {isbn}")
    return ""


def hole_tonie_cover(titel):
    """Sucht Cover auf tonies.club und gibt die finale CDN-URL zurück."""
    try:
        resp = requests.get(
            "https://tonies.club/tonie/all",
            params={"search": titel},
            timeout=10,
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"},
        )
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        img = soup.find("img", class_=lambda c: c and "object-cover" in (c or "") and "h-72" in (c or ""))
        if not img or not img.get("src"):
            print(f"      ⚠️  Kein Tonie-Cover gefunden für '{titel}'")
            return ""
        # Redirect folgen um stabile CDN-URL zu bekommen
        cdn_resp = requests.get(img["src"], timeout=10,
                                headers={"User-Agent": "Mozilla/5.0"}, stream=True)
        cdn_resp.close()
        return cdn_resp.url
    except Exception as exc:
        print(f"      ⚠️  Tonie-Cover fehlgeschlagen für '{titel}': {exc}")
        return ""


def _strip_whitespace(elem):
    """Entfernt reine Whitespace-Textknoten, die minidom beim erneuten Laden multipliziert."""
    for e in elem.iter():
        if e.text and not e.text.strip():
            e.text = None
        if e.tail and not e.tail.strip():
            e.tail = None


def lade_verlauf():
    """Lädt verlauf.xml oder gibt leeres Root-Element zurück."""
    if Path(VERLAUF_XML).exists():
        root = ET.parse(VERLAUF_XML).getroot()
        _strip_whitespace(root)
        return root
    return ET.Element("verlauf")


def speichere_verlauf(root):
    """Speichert verlauf.xml hübsch eingerückt."""
    raw = ET.tostring(root, encoding="unicode")
    dom = minidom.parseString(raw)
    pretty = dom.toprettyxml(indent="  ", encoding="UTF-8")
    with open(VERLAUF_XML, "wb") as f:
        f.write(pretty)


def _hole_cover_fuer_medium(medium_id, titel, mediengruppe, isbn):
    """Wählt die passende Cover-Quelle je nach Mediengruppe."""
    if "Buch" in mediengruppe and isbn:
        time.sleep(2)
        return hole_cover_url(isbn)
    if mediengruppe == "Hörspielzeug" and not titel.startswith("Edurino"):
        return hole_tonie_cover(titel)
    return ""


def aktualisiere_verlauf(alle_medien):
    """Gleicht verlauf.xml mit aktuell ausgeliehenen Medien ab."""
    root = lade_verlauf()
    heute = datetime.date.today().isoformat()
    aktuelle_ids = {m["medium_id"] for m in alle_medien}

    # Rückgaben markieren
    for elem in root.findall("medium"):
        if not _xml_text(elem, "zurueckgegeben") and _xml_text(elem, "medium_id") not in aktuelle_ids:
            elem.find("zurueckgegeben").text = heute
            print(f"   📕 Verlauf: zurückgegeben – {_xml_text(elem, 'titel')}")

    # Neue Medien eintragen / Frist aktualisieren
    for medium in alle_medien:
        mid = medium["medium_id"]
        frist_str = medium["frist"].isoformat()

        # Aktiven Eintrag suchen (noch nicht zurückgegeben)
        aktiv = next(
            (e for e in root.findall("medium")
             if _xml_text(e, "medium_id") == mid and not _xml_text(e, "zurueckgegeben")),
            None
        )

        if aktiv is not None:
            aktiv.find("letzte_frist").text = frist_str
        else:
            isbn = medium.get("isbn", "")
            cover_url = _hole_cover_fuer_medium(mid, medium["titel"], medium.get("mediengruppe", ""), isbn)
            if cover_url:
                print(f"      🖼  Cover gefunden: {medium['titel']}")
            elem = ET.SubElement(root, "medium")
            ET.SubElement(elem, "medium_id").text       = mid
            ET.SubElement(elem, "titel").text            = medium["titel"]
            ET.SubElement(elem, "verfasser").text        = medium.get("verfasser", "")
            ET.SubElement(elem, "mediengruppe").text     = medium.get("mediengruppe", "")
            ET.SubElement(elem, "nutzer").text           = medium["name"]
            ET.SubElement(elem, "isbn").text             = isbn
            ET.SubElement(elem, "cover_url").text        = cover_url
            ET.SubElement(elem, "ausgeliehen_seit").text = heute
            ET.SubElement(elem, "zurueckgegeben").text   = ""
            ET.SubElement(elem, "letzte_frist").text     = frist_str
            print(f"   📗 Verlauf: neu – {medium['titel']}")

    # Cover nachladen für Einträge ohne Cover-URL
    for elem in root.findall("medium"):
        cover_node = elem.find("cover_url")
        if cover_node is None or (cover_node.text or "").strip():
            continue
        mid          = _xml_text(elem, "medium_id")
        titel        = _xml_text(elem, "titel")
        mediengruppe = _xml_text(elem, "mediengruppe")
        isbn         = _xml_text(elem, "isbn")
        url = _hole_cover_fuer_medium(mid, titel, mediengruppe, isbn)
        if url:
            cover_node.text = url
            print(f"      🖼  Cover nachgeladen: {titel}")

    speichere_verlauf(root)
    return root


def generiere_html(root):
    """Generiert verlauf.html aus dem Verlauf-XML-Root."""
    heute = datetime.date.today()

    eintraege = []
    for elem in root.findall("medium"):
        eintraege.append({
            "medium_id":    _xml_text(elem, "medium_id"),
            "titel":        _xml_text(elem, "titel"),
            "verfasser":    _xml_text(elem, "verfasser"),
            "mediengruppe": _xml_text(elem, "mediengruppe"),
            "nutzer":       _xml_text(elem, "nutzer"),
            "isbn":         _xml_text(elem, "isbn"),
            "cover_url":    _xml_text(elem, "cover_url"),
            "seit":         _xml_text(elem, "ausgeliehen_seit"),
            "bis":          _xml_text(elem, "zurueckgegeben"),
            "frist":        _xml_text(elem, "letzte_frist"),
        })

    eintraege.sort(key=lambda x: (x["mediengruppe"].lower(), x["titel"].lower()))

    def fmt_datum(iso):
        try:
            return datetime.date.fromisoformat(iso).strftime("%d.%m.%Y")
        except (ValueError, TypeError):
            return iso

    def tage(seit, bis):
        try:
            d1 = datetime.date.fromisoformat(seit)
            d2 = datetime.date.fromisoformat(bis) if bis else heute
            return (d2 - d1).days
        except (ValueError, TypeError):
            return 0

    nutzer_farben = {"Laura": "#dbeafe", "Benny": "#dcfce7"}

    # Nach Mediengruppe gruppieren
    gruppen = {}
    for e in eintraege:
        gruppen.setdefault(e["mediengruppe"] or "Sonstiges", []).append(e)

    def karte_html(e):
        ist_aktiv = not e["bis"]
        farbe = nutzer_farben.get(e["nutzer"], "#f3f4f6")
        t = tage(e["seit"], e["bis"])

        if e["cover_url"]:
            cover_img = (
                f'<img src="{e["cover_url"]}" alt="" loading="lazy" '
                f'onerror="this.style.display=\'none\';this.nextElementSibling.style.display=\'flex\'">'
                f'<div class="no-cover" style="display:none">📚</div>'
            )
        else:
            cover_img = '<div class="no-cover">📚</div>'

        dot = '<span class="dot-aktiv"></span>' if ist_aktiv else ''
        verfasser_html = f'<div class="verfasser">{e["verfasser"]}</div>' if e["verfasser"] else ""
        datum_zeile = fmt_datum(e["seit"])
        if e["bis"]:
            datum_zeile += f' – {fmt_datum(e["bis"])}'

        return (
            f'<div class="karte{"" if ist_aktiv else " karte-zurueck"}">'
            f'<a class="cover-wrap" href="https://katalog.halle.de/Mediensuche?id={e["medium_id"]}" target="_blank">'
            f'{cover_img}'
            f'{dot}'
            f'<span class="nutzer-badge" style="background:{farbe}">{e["nutzer"]}</span>'
            f'</a>'
            f'<div class="info">'
            f'<div class="titel">{e["titel"]}</div>'
            f'{verfasser_html}'
            f'<div class="datum">📅 {datum_zeile} · {t} {"Tag" if t == 1 else "Tage"}</div>'
            f'</div>'
            f'</div>'
        )

    sektionen = []
    for gruppe, items in sorted(gruppen.items(), key=lambda g: g[0].lower()):
        anz = len(items)
        anz_aktiv_gruppe = sum(1 for e in items if not e["bis"])
        aktiv_hint = f" · {anz_aktiv_gruppe} ausgeliehen" if anz_aktiv_gruppe else ""
        karten_block = "\n".join(karte_html(e) for e in items)
        sektionen.append(
            f'<details open>\n'
            f'  <summary>{gruppe} <span class="gruppe-anzahl">({anz}{aktiv_hint})</span></summary>\n'
            f'  <div class="raster">\n{karten_block}\n  </div>\n'
            f'</details>'
        )

    anz_gesamt = len(eintraege)
    anz_aktiv  = sum(1 for e in eintraege if not e["bis"])
    generiert  = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    sektionen_html = "\n".join(sektionen)

    if SEITEN_PASSWORT:
        pw_hash = hashlib.sha256(SEITEN_PASSWORT.encode()).hexdigest()
        pw_overlay_html = f"""<div id="pw-overlay">
  <div id="pw-box">
    <h2>📚 Bibliothek</h2>
    <p>Bitte Passwort eingeben</p>
    <input id="pw-input" type="password" placeholder="Passwort" autofocus
           onkeydown="if(event.key==='Enter')checkPw()">
    <button id="pw-btn" onclick="checkPw()">Weiter</button>
    <div id="pw-error">Falsches Passwort</div>
  </div>
</div>"""
        pw_script_html = f"""<script>
  (async () => {{
    if (sessionStorage.getItem('bib_auth') === '1')
      document.getElementById('pw-overlay').style.display = 'none';
  }})();
  async function checkPw() {{
    const val = document.getElementById('pw-input').value;
    const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(val));
    const hex = Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,'0')).join('');
    if (hex === '{pw_hash}') {{
      sessionStorage.setItem('bib_auth', '1');
      document.getElementById('pw-overlay').style.display = 'none';
    }} else {{
      document.getElementById('pw-error').style.display = 'block';
      document.getElementById('pw-input').value = '';
    }}
  }}
</script>"""
    else:
        pw_overlay_html = ""
        pw_script_html  = ""

    html = f"""<!DOCTYPE html>
<html lang="de">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Bibliothek Verlauf</title>
  <style>
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      background: #f0f2f5;
      color: #1f2937;
      padding: 24px 20px;
      min-height: 100vh;
    }}
    header {{ margin-bottom: 24px; }}
    h1 {{ font-size: 1.75rem; color: #111827; margin-bottom: 4px; }}
    .subtitle {{ font-size: 0.875rem; color: #6b7280; }}
    details {{ margin-bottom: 20px; }}
    details[open] summary {{ border-radius: 10px 10px 0 0; }}
    summary {{
      list-style: none;
      cursor: pointer;
      background: #fff;
      border-radius: 10px;
      padding: 12px 16px;
      font-size: 1rem;
      font-weight: 700;
      color: #111827;
      box-shadow: 0 1px 3px rgba(0,0,0,.08);
      display: flex;
      align-items: center;
      gap: 8px;
      user-select: none;
    }}
    summary::before {{
      content: "▶";
      font-size: 0.65rem;
      color: #9ca3af;
      transition: transform .2s;
      flex-shrink: 0;
    }}
    details[open] summary::before {{ transform: rotate(90deg); }}
    summary::-webkit-details-marker {{ display: none; }}
    .gruppe-anzahl {{ font-size: 0.8rem; font-weight: 400; color: #6b7280; }}
    .raster {{
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(150px, 1fr));
      gap: 12px;
      background: #fff;
      border-radius: 0 0 10px 10px;
      padding: 14px;
      box-shadow: 0 1px 3px rgba(0,0,0,.08);
    }}
    .karte {{
      background: #f9fafb;
      border-radius: 10px;
      overflow: hidden;
      display: flex;
      flex-direction: column;
      box-shadow: 0 1px 4px rgba(0,0,0,.1);
    }}
    .karte-zurueck {{ opacity: 0.6; box-shadow: none; }}
    .cover-wrap {{
      position: relative;
      width: 100%;
      aspect-ratio: 2 / 3;
      background: #e5e7eb;
      display: block;
      overflow: hidden;
      text-decoration: none;
    }}
    .cover-wrap img {{
      width: 100%;
      height: 100%;
      object-fit: cover;
      display: block;
    }}
    .no-cover {{
      width: 100%;
      height: 100%;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 2.5rem;
      background: #e5e7eb;
    }}
    .dot-aktiv {{
      position: absolute;
      top: 8px;
      right: 8px;
      width: 13px;
      height: 13px;
      background: #4ade80;
      border-radius: 50%;
      box-shadow: 0 0 0 2px rgba(255,255,255,0.85);
    }}
    .nutzer-badge {{
      position: absolute;
      bottom: 7px;
      right: 7px;
      font-size: 0.68rem;
      font-weight: 600;
      padding: 2px 8px;
      border-radius: 999px;
      border: 1px solid rgba(0,0,0,.08);
    }}
    .info {{
      padding: 9px 10px 11px;
      display: flex;
      flex-direction: column;
      gap: 3px;
    }}
    .titel {{
      font-weight: 700;
      font-size: 0.82rem;
      line-height: 1.3;
      color: #111827;
    }}
    .verfasser {{ font-size: 0.74rem; color: #6b7280; }}
    .datum {{ font-size: 0.7rem; color: #9ca3af; margin-top: 2px; }}
    footer {{ margin-top: 32px; text-align: center; font-size: 0.78rem; color: #9ca3af; }}
    #pw-overlay {{
      position: fixed; inset: 0;
      background: #f0f2f5;
      display: flex; align-items: center; justify-content: center;
      z-index: 9999;
    }}
    #pw-box {{
      background: #fff;
      border-radius: 16px;
      padding: 36px 32px;
      box-shadow: 0 4px 24px rgba(0,0,0,.12);
      text-align: center;
      width: 320px;
    }}
    #pw-box h2 {{ font-size: 1.4rem; margin-bottom: 6px; }}
    #pw-box p {{ font-size: 0.85rem; color: #6b7280; margin-bottom: 20px; }}
    #pw-input {{
      width: 100%; padding: 10px 14px;
      border: 1.5px solid #e5e7eb; border-radius: 8px;
      font-size: 1rem; outline: none;
      transition: border-color .15s;
    }}
    #pw-input:focus {{ border-color: #2563eb; }}
    #pw-btn {{
      margin-top: 12px; width: 100%;
      padding: 10px; border: none; border-radius: 8px;
      background: #2563eb; color: #fff;
      font-size: 0.95rem; font-weight: 600; cursor: pointer;
    }}
    #pw-btn:hover {{ background: #1d4ed8; }}
    #pw-error {{ color: #dc2626; font-size: 0.82rem; margin-top: 10px; display: none; }}
  </style>
</head>
<body>
  {pw_overlay_html}
  <header>
    <h1>📚 Bibliothek Verlauf</h1>
    <div class="subtitle">{anz_gesamt} Medien insgesamt · {anz_aktiv} aktuell ausgeliehen · Stand {generiert}</div>
  </header>
  {sektionen_html}
  <footer>Stadtbibliothek Halle · generiert von bibliothek_kalender_sync.py</footer>
  {pw_script_html}
</body>
</html>"""

    with open(VERLAUF_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"   📄 verlauf.html generiert ({anz_gesamt} Einträge, {len(gruppen)} Gruppen)")


# ─────────────────────────────────────────────
# SCHRITT 4: Hauptprogramm
# ─────────────────────────────────────────────

def main():
    print("\n" + "═" * 55)
    print("  📚 Bibliothek Halle → Google Calendar Sync")
    print("═" * 55 + "\n")

    # 1. Medien aller Benutzer abrufen
    alle_medien = []
    for benutzer in BENUTZER:
        print(f"─── Ausgeliehene Medien: {benutzer['name']} ───")
        try:
            medien = hole_ausgeliehene_medien(
                benutzer["name"], benutzer["ausweis"], benutzer["passwort"]
            )
            alle_medien.extend(medien)
        except Exception as e:
            print(f"\n❌ Fehler beim Abruf für {benutzer['name']}:\n   {e}")

    if not alle_medien:
        print("ℹ️  Aktuell keine Medien ausgeliehen.")

    # 2. Nach Datum gruppieren
    medien_nach_datum = {}
    for medium in alle_medien:
        medien_nach_datum.setdefault(medium["frist"].isoformat(), []).append(medium)

    # 3. Google Calendar verbinden
    print("\n─── Google Calendar ───")
    try:
        service = google_calendar_service()
    except Exception as e:
        print(f"\n❌ Fehler bei Google-Authentifizierung:\n   {e}")
        return

    # 4. Bestehende Events laden
    bestehende_events, alte_events = hole_bestehende_biblio_events(service)
    print(f"   Gefunden: {len(bestehende_events)} bestehende Einträge im Kalender.")
    if alte_events:
        print(f"   Migration: {len(alte_events)} Einträge im alten Format werden gelöscht.")

    # 4a. Alte Events (ein Event pro Medium) löschen
    for event in alte_events:
        service.events().delete(calendarId=KALENDER_ID, eventId=event["id"]).execute()
        print(f"   🗑️  Migration: altes Event gelöscht: {event.get('summary', event['id'])}")

    # 5. Events synchronisieren
    print("\n─── Kalender aktualisieren ───")
    sync_events(service, medien_nach_datum, bestehende_events)

    # 6. Verlauf aktualisieren und HTML generieren
    print("\n─── Verlauf aktualisieren ───")
    verlauf_root = aktualisiere_verlauf(alle_medien)
    generiere_html(verlauf_root)

    print("\n" + "═" * 55)
    print("  ✅ Sync abgeschlossen!")
    print(f"     Erinnerung: {ERINNERUNG_TAGE_VORHER} Tage vor Ablauf")
    print("═" * 55 + "\n")


# ─────────────────────────────────────────────
# DIAGNOSE-FUNKTION (bei Login-Problemen)
# ─────────────────────────────────────────────

def run_debug_login():
    """
    Hilfsfunktion zur Diagnose: Zeigt alle Formularfelder der Login-Seite.
    Aufrufen mit: python bibliothek_kalender_sync.py --debug
    """
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    response = session.get(LOGIN_URL)
    soup = BeautifulSoup(response.text, "html.parser")

    print("\n=== Formularfelder der Kontoseite ===")
    for inp in soup.find_all("input"):
        typ = inp.get("type", "text")
        if typ != "hidden":
            print(f"  [{typ}] name='{inp.get('name', '')}'")

    print("\n=== Versteckte Felder ===")
    for inp in soup.find_all("input", {"type": "hidden"}):
        print(f"  name='{inp.get('name')}' value='{inp.get('value','')[:40]}...'")


# ─────────────────────────────────────────────
# EINSTIEGSPUNKT
# ─────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    if "--debug" in sys.argv:
        run_debug_login()
    else:
        main()


# ══════════════════════════════════════════════════════════════════
# README / SETUP-ANLEITUNG
# ══════════════════════════════════════════════════════════════════
#
# VORAUSSETZUNGEN
# ───────────────
# Python 3.8+, dann einmalig installieren:
#
#   pip install requests beautifulsoup4 google-auth google-auth-oauthlib google-api-python-client
#
#
# SCHRITT 1: Bibliotheksdaten eintragen
# ─────────────────────────────────────
# Umgebungsvariablen setzen (z.B. als GitHub Secrets):
#
#   BIBL_AUSWEIS_LAURA   – Ausweisnummer Laura
#   BIBL_PASSWORT_LAURA  – Passwort Laura
#   BIBL_AUSWEIS_BENNY   – Ausweisnummer Benny
#   BIBL_PASSWORT_BENNY  – Passwort Benny
#
# Weitere Benutzer können in der BENUTZER-Liste oben hinzugefügt werden.
#
#
# SCHRITT 2: Google Calendar API einrichten
# ─────────────────────────────────────────
# 1. Gehe zu: https://console.cloud.google.com/
# 2. Neues Projekt erstellen (z.B. "Bibliothek Sync")
# 3. "APIs & Dienste" → "Bibliothek" → "Google Calendar API" aktivieren
# 4. "APIs & Dienste" → "Anmeldedaten" → "+ Anmeldedaten erstellen"
#    → "OAuth-Client-ID" → Anwendungstyp: "Desktop-App"
# 5. JSON herunterladen, als "credentials.json" neben dieses Script legen
# 6. Beim ersten Start öffnet sich ein Browser → Google-Konto auswählen
#    → Zugriff erlauben. Danach wird token.json automatisch erstellt.
#
#
# SCHRITT 3: Eigenen Kalender anlegen (empfohlen)
# ────────────────────────────────────────────────
# In Google Calendar: + Neuer Kalender → z.B. "Bibliothek Fristen"
# Kalender-ID findest du unter: Einstellungen → Kalender → Kalender-ID
# Diese ID oben als KALENDER_ID eintragen.
#
#
# SCHRITT 4: Regelmäßig ausführen mit GitHub Actions
# ────────────────────────────────────────────────────
# Siehe beiliegende Datei: .github/workflows/sync.yml
# Die Zugangsdaten werden als GitHub Secrets hinterlegt (niemals im Code!).
#
#
# MANUELL AUSFÜHREN
# ─────────────────
#   python bibliothek_kalender_sync.py
#
# BEI LOGIN-PROBLEMEN
# ───────────────────
#   python bibliothek_kalender_sync.py --debug
#
