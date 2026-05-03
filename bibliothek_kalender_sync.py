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
import datetime
import xml.etree.ElementTree as ET
from collections import defaultdict
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

# ID des Ziel-Kalenders – aus Umgebungsvariable (GitHub Secret: KALENDER_ID)
KALENDER_ID = os.environ.get("KALENDER_ID", "")

# Wie viele Tage vor Ablauf soll der Erinnerungsalarm erscheinen?
ERINNERUNG_TAGE_VORHER = 2

# Medien mit Frist innerhalb dieser Tage werden automatisch verlängert
VERLAENGERUNG_SCHWELLE_TAGE = 3

# Verlauf-Dateien
VERLAUF_XML        = "verlauf.xml"
VERLAUF_HTML       = "verlauf.html"
VERLAUF_HTML_LOKAL = "verlauf_lokal.html"
VERLAUF_STATISTIK  = "statistik.html"
COVERS_DIR         = "covers"

# Optionaler Passwortschutz für verlauf.html (leer = kein Schutz)
SEITEN_PASSWORT = os.environ.get("HTML_PASSWORT", "")

# ─────────────────────────────────────────────
# KONSTANTEN
# ─────────────────────────────────────────────

LOGIN_URL = "https://katalog.halle.de/Mein-Konto"
SCOPES    = ["https://www.googleapis.com/auth/calendar"]


def _get_hidden(soup, field_name):
    tag = soup.find("input", {"name": field_name})
    return tag["value"] if tag else ""


def _ist_edurino(titel):
    return titel.startswith("Edurino")

NUTZER_FARBEN_VERLAUF    = {"Laura": "#dbeafe", "Benny": "#dcfce7"}
NUTZER_FARBEN_STATISTIK  = {"Laura": "#3b82f6", "Benny": "#22c55e"}


# ─────────────────────────────────────────────
# SCHRITT 1: Bibliothekskonto scrapen
# ─────────────────────────────────────────────

def verlaengere_faellige_medien(session, soup, name):
    """
    Verlängert alle verlängerbaren Medien mit Frist ≤ heute + VERLAENGERUNG_SCHWELLE_TAGE.
    Gibt das aktualisierte HTML zurück, oder None wenn nichts zu verlängern war.
    """
    heute = datetime.date.today()
    grenz_datum = heute + datetime.timedelta(days=VERLAENGERUNG_SCHWELLE_TAGE)

    tabelle = soup.find("table", {"id": lambda x: x and "grdViewLoans" in x})
    if not tabelle:
        return None

    # Schritt 1: Fällige, verlängerbare Medien und ihre Checkbox-Namen ermitteln
    zu_verlaengern = []

    for zeile in tabelle.find_all("tr")[1:]:
        zellen = zeile.find_all("td")
        if len(zellen) < 4:
            continue

        # Aktuelles Fristdatum aus dem "Aktuelle Frist"-Label lesen
        datum = None
        for zelle in zellen:
            label = zelle.find("span", string=re.compile(r"Aktuelle Frist"))
            if label is None:
                label = zelle.find("span", class_=lambda c: c and "oclc-module-label" in c)
                if label and "Aktuelle Frist" not in label.get_text():
                    label = None
            if label:
                frist_text = zelle.get_text(strip=True)
                m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", frist_text)
                if m:
                    datum = datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
                break

        if datum is None or datum > grenz_datum:
            continue

        checkbox_span = zeile.find("span", class_=lambda c: c and "loancheckbox" in c)
        checkbox = checkbox_span.find("input", {"type": "checkbox"}) if checkbox_span else None
        if not checkbox:
            continue

        titel_link = zeile.find("a", href=lambda h: h and "Mediensuche" in h)
        titel = titel_link.get_text(strip=True) if titel_link else "?"
        zu_verlaengern.append({"titel": titel, "frist": datum, "checkbox_name": checkbox["name"]})

    if not zu_verlaengern:
        return None

    print(f"\n   🔄 {len(zu_verlaengern)} Medium/Medien werden verlängert ({name}):")
    for item in zu_verlaengern:
        print(f"      • {item['titel']} (Frist: {item['frist'].strftime('%d.%m.%Y')})")

    # Schritt 2: Alle hidden-Felder der Seite einsammeln (wie ein Browser)
    post_data = {}
    for inp in soup.find_all("input", type="hidden"):
        n = inp.get("name")
        if n:
            post_data[n] = inp.get("value", "")

    # Schritt 3: Checkboxen der zu verlängernden Medien aktivieren
    for item in zu_verlaengern:
        post_data[item["checkbox_name"]] = "on"

    # Schritt 4: Submit-Knopf "Medien verlängern" hinzufügen
    post_data["dnn$ctr376$MainView$tpnlLoans$ucLoansView$BtnExtendMediums"] = "Medien verlängern"

    response = session.post(LOGIN_URL, data=post_data, timeout=20)
    response.raise_for_status()
    print(f"   ✅ Verlängerung abgeschickt ({name}).")
    return response.text


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
    response = session.get(LOGIN_URL, timeout=20)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")
    
    post_data = {
        "__EVENTTARGET":        "",
        "__EVENTARGUMENT":      "",
        "__VIEWSTATE":          _get_hidden(soup, "__VIEWSTATE"),
        "__VIEWSTATEGENERATOR": _get_hidden(soup, "__VIEWSTATEGENERATOR"),
        "__EVENTVALIDATION":    _get_hidden(soup, "__EVENTVALIDATION"),
        "__VIEWSTATEENCRYPTED": _get_hidden(soup, "__VIEWSTATEENCRYPTED"),
        "dnn$ctr375$Login$Login_COP$txtUsername": ausweis,
        "dnn$ctr375$Login$Login_COP$txtPassword": passwort,
        "dnn$ctr375$Login$Login_COP$cmdLogin":    "Anmelden",
    }

    print(f"🔐 Melde an ({name})...")
    response = session.post(LOGIN_URL, data=post_data, timeout=20)
    response.raise_for_status()

    if "abmelden" not in response.text.lower():
        raise RuntimeError(
            f"❌ Login fehlgeschlagen für {name}. Bitte Ausweisnummer und Passwort prüfen.\n"
            "   Hinweis: run_debug_login() aufrufen für Diagnose."
        )
    print(f"✅ Login erfolgreich ({name}).")

    login_soup = BeautifulSoup(response.text, "html.parser")

    try:
        aktualisiertes_html = verlaengere_faellige_medien(session, login_soup, name)
        html = aktualisiertes_html if aktualisiertes_html else response.text
    except Exception as e:
        print(f"   ⚠️  Verlängerung fehlgeschlagen ({name}), Sync wird fortgesetzt: {e}")
        html = response.text
    return parse_ausleih_seite(html, name)


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


def _medien_prefix(m):
    mg = m.get("mediengruppe", "")
    if mg == "Hörspielzeug":
        return "" if _ist_edurino(m.get("titel", "")) else "Tonie: "
    if mg == "Kinder-CD":
        return "CD: "
    return ""


def baue_event_body(datum_str, medien_gruppe):
    """Erstellt den Event-Body für eine Gruppe von Medien am gleichen Tag."""
    anzahl = len(medien_gruppe)
    summary = "📚 1 Medium abgeben" if anzahl == 1 else f"📚 {anzahl} Medien abgeben"

    beschreibung = "\n".join(
        f"- {_medien_prefix(m)}{m['titel']} ({m['name']})"
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

            alte_beschreibung = altes_event.get("description", "")
            if alte_ids == neue_ids and alte_beschreibung == neues_body["description"]:
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


def _aktive_ausleihe(medium_elem):
    """Gibt die aktive <ausleihe> zurück (ohne <zurueck>-Datum), oder None."""
    ausleihen = medium_elem.find("ausleihen")
    if ausleihen is None:
        return None
    for a in ausleihen.findall("ausleihe"):
        if not (_xml_text(a, "zurueck")).strip():
            return a
    return None


def _migriere_wenn_noetig(root):
    """Konvertiert altes Flat-Format (ausgeliehen_seit/zurueckgegeben/letzte_frist) zum neuen <ausleihen>-Format."""
    for elem in root.findall("medium"):
        if elem.find("ausgeliehen_seit") is None:
            continue
        nutzer  = _xml_text(elem, "nutzer")
        seit    = _xml_text(elem, "ausgeliehen_seit")
        zurueck = _xml_text(elem, "zurueckgegeben")
        frist   = _xml_text(elem, "letzte_frist")
        for tag in ("ausgeliehen_seit", "zurueckgegeben", "letzte_frist"):
            old = elem.find(tag)
            if old is not None:
                elem.remove(old)
        ausleihen = ET.SubElement(elem, "ausleihen")
        ausleihe  = ET.SubElement(ausleihen, "ausleihe")
        ET.SubElement(ausleihe, "nutzer").text  = nutzer
        ET.SubElement(ausleihe, "seit").text    = seit
        ET.SubElement(ausleihe, "zurueck").text = zurueck
        ET.SubElement(ausleihe, "frist").text   = frist


def _migriere_nutzer_in_ausleihe(root):
    """Fügt <nutzer> in bestehende <ausleihe>-Einträge ein, die ihn noch nicht haben."""
    for elem in root.findall("medium"):
        nutzer = _xml_text(elem, "nutzer")
        ausleihen = elem.find("ausleihen")
        if ausleihen is None:
            continue
        for a in ausleihen.findall("ausleihe"):
            if a.find("nutzer") is None:
                nutzer_elem = ET.Element("nutzer")
                nutzer_elem.text = nutzer
                a.insert(0, nutzer_elem)


def hole_cover_url(isbn):
    """Sucht Cover: DNB (1), Open Library (2), Google Books (3)."""
    if not isbn:
        return ""

    # 1. DNB
    dnb_url = f"https://portal.dnb.de/opac/mvb/cover?isbn={isbn}"
    try:
        resp = requests.head(dnb_url, timeout=8, allow_redirects=True)
        if resp.status_code == 200 and "image" in resp.headers.get("Content-Type", ""):
            return dnb_url
    except Exception as exc:
        print(f"      ⚠️  DNB fehlgeschlagen für ISBN {isbn}: {exc}")

    # 2. Open Library
    ol_url = f"https://covers.openlibrary.org/b/isbn/{isbn}-L.jpg"
    try:
        resp = requests.head(ol_url, params={"default": "false"}, timeout=8, allow_redirects=True)
        if resp.status_code == 200:
            return ol_url
    except Exception as exc:
        print(f"      ⚠️  Open Library fehlgeschlagen für ISBN {isbn}: {exc}")

    # 3. Google Books
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
        _migriere_wenn_noetig(root)
        _migriere_nutzer_in_ausleihe(root)
        return root
    return ET.Element("verlauf")


def speichere_verlauf(root):
    """Speichert verlauf.xml hübsch eingerückt."""
    raw = ET.tostring(root, encoding="unicode")
    dom = minidom.parseString(raw)
    pretty = dom.toprettyxml(indent="  ", encoding="UTF-8")
    with open(VERLAUF_XML, "wb") as f:
        f.write(pretty)


def lade_cover_lokal(cover_url, medium_id):
    """Lädt Cover von cover_url in COVERS_DIR herunter (gecacht)."""
    if not cover_url:
        return ""
    Path(COVERS_DIR).mkdir(exist_ok=True)
    local_path = Path(COVERS_DIR) / f"{medium_id}.jpg"
    if local_path.exists():
        return str(local_path).replace("\\", "/")
    try:
        resp = requests.get(cover_url, timeout=10, headers={"User-Agent": "Mozilla/5.0"})
        resp.raise_for_status()
        local_path.write_bytes(resp.content)
        return str(local_path).replace("\\", "/")
    except Exception as exc:
        print(f"      ⚠️  Cover-Download fehlgeschlagen für {medium_id}: {exc}")
        return ""


def _hole_cover_fuer_medium(medium_id, titel, mediengruppe, isbn):
    """Wählt die passende Cover-Quelle je nach Mediengruppe."""
    if isbn:
        time.sleep(2)
        return hole_cover_url(isbn)
    if mediengruppe == "Hörspielzeug" and not _ist_edurino(titel):
        return hole_tonie_cover(titel)
    return ""


def aktualisiere_verlauf(alle_medien):
    """Gleicht verlauf.xml mit aktuell ausgeliehenen Medien ab."""
    root = lade_verlauf()
    heute = datetime.date.today().isoformat()
    aktuelle_ids = {m["medium_id"] for m in alle_medien}

    # Rückgaben markieren
    for elem in root.findall("medium"):
        if _xml_text(elem, "medium_id") not in aktuelle_ids:
            aktiv = _aktive_ausleihe(elem)
            if aktiv is not None:
                aktiv.find("zurueck").text = heute
                print(f"   📕 Verlauf: zurückgegeben – {_xml_text(elem, 'titel')}")

    # Neue Medien eintragen / Frist aktualisieren
    for medium in alle_medien:
        mid = medium["medium_id"]
        frist_str = medium["frist"].isoformat()

        medium_elem = next(
            (e for e in root.findall("medium") if _xml_text(e, "medium_id") == mid),
            None
        )

        if medium_elem is not None:
            aktiv = _aktive_ausleihe(medium_elem)
            if aktiv is not None:
                aktiv.find("frist").text = frist_str
            else:
                # Erneute Ausleihe desselben Mediums
                ausleihen = medium_elem.find("ausleihen")
                ausleihe  = ET.SubElement(ausleihen, "ausleihe")
                ET.SubElement(ausleihe, "nutzer").text  = medium["name"]
                ET.SubElement(ausleihe, "seit").text    = heute
                ET.SubElement(ausleihe, "zurueck").text = ""
                ET.SubElement(ausleihe, "frist").text   = frist_str
                print(f"   📗 Verlauf: erneut ausgeliehen – {_xml_text(medium_elem, 'titel')}")
        else:
            isbn = medium.get("isbn", "")
            local_path = Path(COVERS_DIR) / f"{mid}.jpg"
            if local_path.exists():
                cover_url   = ""
                cover_lokal = str(local_path).replace("\\", "/")
            else:
                cover_url = _hole_cover_fuer_medium(mid, medium["titel"], medium.get("mediengruppe", ""), isbn)
                if cover_url:
                    print(f"      🖼  Cover gefunden: {medium['titel']}")
                cover_lokal = lade_cover_lokal(cover_url, mid)
            elem = ET.SubElement(root, "medium")
            ET.SubElement(elem, "medium_id").text   = mid
            ET.SubElement(elem, "titel").text        = medium["titel"]
            ET.SubElement(elem, "verfasser").text    = medium.get("verfasser", "")
            ET.SubElement(elem, "mediengruppe").text = medium.get("mediengruppe", "")
            ET.SubElement(elem, "nutzer").text       = medium["name"]
            ET.SubElement(elem, "isbn").text         = isbn
            ET.SubElement(elem, "cover_url").text    = cover_url
            ET.SubElement(elem, "cover_lokal").text  = cover_lokal
            ausleihen = ET.SubElement(elem, "ausleihen")
            ausleihe  = ET.SubElement(ausleihen, "ausleihe")
            ET.SubElement(ausleihe, "nutzer").text  = medium["name"]
            ET.SubElement(ausleihe, "seit").text    = heute
            ET.SubElement(ausleihe, "zurueck").text = ""
            ET.SubElement(ausleihe, "frist").text   = frist_str
            print(f"   📗 Verlauf: neu – {medium['titel']}")

    # Cover nachladen für Einträge ohne cover_url oder cover_lokal
    for elem in root.findall("medium"):
        mid          = _xml_text(elem, "medium_id")
        titel        = _xml_text(elem, "titel")
        mediengruppe = _xml_text(elem, "mediengruppe")
        isbn         = _xml_text(elem, "isbn")

        cover_node = elem.find("cover_url")
        if cover_node is None:
            cover_node = ET.SubElement(elem, "cover_url")

        cover_lokal_node = elem.find("cover_lokal")
        if cover_lokal_node is None:
            cover_lokal_node = ET.SubElement(elem, "cover_lokal")

        local_path = Path(COVERS_DIR) / f"{mid}.jpg"
        if local_path.exists():
            if not (cover_lokal_node.text or "").strip():
                cover_lokal_node.text = str(local_path).replace("\\", "/")
            continue

        if not (cover_node.text or "").strip():
            url = _hole_cover_fuer_medium(mid, titel, mediengruppe, isbn)
            if url:
                cover_node.text = url
                print(f"      🖼  Cover nachgeladen: {titel}")

        if not (cover_lokal_node.text or "").strip() and (cover_node.text or "").strip():
            path = lade_cover_lokal(cover_node.text, mid)
            if path:
                cover_lokal_node.text = path
                print(f"      🖼  Cover lokal gespeichert: {titel}")

    speichere_verlauf(root)
    return root


def generiere_html(root, lokal=False):
    """Generiert verlauf.html (lokal=False) oder verlauf_lokal.html (lokal=True)."""
    heute = datetime.date.today()

    eintraege = []
    for elem in root.findall("medium"):
        ausleihen_elem = elem.find("ausleihen")
        if ausleihen_elem is None:
            continue
        alle_ausleihen = ausleihen_elem.findall("ausleihe")
        if not alle_ausleihen:
            continue
        aktive = next((a for a in alle_ausleihen if not _xml_text(a, "zurueck").strip()), None)
        referenz = aktive if aktive is not None else alle_ausleihen[-1]
        eintraege.append({
            "medium_id":    _xml_text(elem, "medium_id"),
            "titel":        _xml_text(elem, "titel"),
            "verfasser":    _xml_text(elem, "verfasser"),
            "mediengruppe": _xml_text(elem, "mediengruppe"),
            "nutzer":       _xml_text(elem, "nutzer"),
            "isbn":         _xml_text(elem, "isbn"),
            "cover_url":    _xml_text(elem, "cover_url"),
            "cover_lokal":  _xml_text(elem, "cover_lokal"),
            "seit":         _xml_text(referenz, "seit"),
            "bis":          _xml_text(referenz, "zurueck"),
            "frist":        _xml_text(referenz, "frist"),
            "ausleihen":    [
                {"nutzer": _xml_text(a, "nutzer"), "seit": _xml_text(a, "seit"), "zurueck": _xml_text(a, "zurueck"), "frist": _xml_text(a, "frist")}
                for a in alle_ausleihen
            ],
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

    def display_gruppe(e):
        if e["mediengruppe"] == "Hörspielzeug":
            return "Edurinos" if _ist_edurino(e["titel"]) else "Tonies"
        return e["mediengruppe"] or "Sonstiges"

    # Nach Mediengruppe gruppieren
    gruppen = {}
    for e in eintraege:
        gruppen.setdefault(display_gruppe(e), []).append(e)

    def karte_html(e):
        ist_aktiv = not e["bis"]
        farbe = NUTZER_FARBEN_VERLAUF.get(e["nutzer"], "#f3f4f6")
        t = tage(e["seit"], e["bis"])

        cover_src = (e["cover_lokal"] or e["cover_url"]) if lokal else e["cover_url"]
        if cover_src:
            cover_img = (
                f'<img src="{cover_src}" alt="" loading="lazy" '
                f'onerror="this.style.display=\'none\';this.nextElementSibling.style.display=\'flex\'">'
                f'<div class="no-cover" style="display:none">📚</div>'
            )
        else:
            cover_img = '<div class="no-cover">📚</div>'

        dot = '<span class="dot-aktiv"></span>' if ist_aktiv else ''
        verfasser_html = f'<div class="verfasser">{e["verfasser"]}</div>' if e["verfasser"] else ""
        def fmt_ausleihe_zeile(a, aktuell):
            zeile = fmt_datum(a["seit"])
            if a["zurueck"]:
                zeile += f' – {fmt_datum(a["zurueck"])}'
                zusatz = f' · {tage(a["seit"], a["zurueck"])} {"Tag" if tage(a["seit"], a["zurueck"]) == 1 else "Tage"}'
            elif a["frist"]:
                zusatz = f' bis {fmt_datum(a["frist"])}'
            else:
                zusatz = ''
            css = "datum" if aktuell else "datum datum-alt"
            return f'<div class="{css}">📅 {zeile}{zusatz}</div>'

        ausleihen_html = "".join(
            fmt_ausleihe_zeile(a, i == len(e["ausleihen"]) - 1)
            for i, a in enumerate(e["ausleihen"])
        )

        return (
            f'<div class="karte{"" if ist_aktiv else " karte-zurueck"}" data-nutzer="{e["nutzer"].lower()}">'
            f'<a class="cover-wrap" href="https://katalog.halle.de/Mediensuche?id={e["medium_id"]}" target="_blank">'
            f'{cover_img}'
            f'{dot}'
            f'<span class="nutzer-badge" style="background:{farbe}">{e["nutzer"]}</span>'
            f'</a>'
            f'<div class="info">'
            f'<div class="titel">{e["titel"]}</div>'
            f'{verfasser_html}'
            f'{ausleihen_html}'
            f'</div>'
            f'</div>'
        )

    gruppen_reihenfolge = ["Kinder-Buch", "Tonies", "Edurinos", "Kinder-CD", "Spiel"]

    def gruppe_sort_key(g):
        try:
            return gruppen_reihenfolge.index(g[0])
        except ValueError:
            return len(gruppen_reihenfolge)

    sektionen = []
    for gruppe, items in sorted(gruppen.items(), key=gruppe_sort_key):
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
    alle_nutzer = sorted({e["nutzer"] for e in eintraege})
    sektionen_html = "\n".join(sektionen)

    if SEITEN_PASSWORT:
        # FNV-1a 32-bit – läuft überall (kein crypto.subtle nötig)
        def _fnv1a(s):
            h = 0x811c9dc5
            for c in s.encode("utf-8"):
                h = ((h ^ c) * 0x01000193) & 0xFFFFFFFF
            return format(h, "08x")

        pw_hash = _fnv1a(SEITEN_PASSWORT)
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
  if (sessionStorage.getItem('bib_auth') === '1')
    document.getElementById('pw-overlay').style.display = 'none';
  function checkPw() {{
    const val = document.getElementById('pw-input').value;
    let h = 0x811c9dc5;
    for (let i = 0; i < val.length; i++)
      h = Math.imul(h ^ val.charCodeAt(i), 0x01000193) >>> 0;
    const hex = h.toString(16).padStart(8, '0');
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
    .datum-alt {{ font-size: 0.65rem; color: #d1d5db; }}
    footer {{ margin-top: 32px; text-align: center; font-size: 0.78rem; color: #9ca3af; }}
    #aktiv-filter-btn {{
      padding: 6px 14px;
      border: 1.5px solid #d1d5db;
      border-radius: 999px;
      background: #fff;
      color: #374151;
      font-size: 0.82rem;
      font-weight: 600;
      cursor: pointer;
      transition: background .15s, border-color .15s, color .15s;
    }}
    #aktiv-filter-btn:hover {{ background: #f3f4f6; }}
    #aktiv-filter-btn.aktiv {{
      background: #2563eb;
      border-color: #2563eb;
      color: #fff;
    }}
    #aktiv-filter-btn.aktiv:hover {{ background: #1d4ed8; }}
    .nutzer-filter-btn {{
      padding: 6px 14px;
      border: 1.5px solid #d1d5db;
      border-radius: 999px;
      background: #fff;
      color: #374151;
      font-size: 0.82rem;
      font-weight: 600;
      cursor: pointer;
      transition: background .15s, border-color .15s, color .15s;
    }}
    .nutzer-filter-btn:hover {{ background: #f3f4f6; }}
    .nutzer-filter-btn.aktiv {{
      background: #374151;
      border-color: #374151;
      color: #fff;
    }}
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
    #suche {{
      width: 100%;
      max-width: 420px;
      padding: 8px 14px 8px 36px;
      border: 1.5px solid #d1d5db;
      border-radius: 999px;
      font-size: 0.9rem;
      outline: none;
      background: #fff url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='16' height='16' fill='none' stroke='%239ca3af' stroke-width='2' viewBox='0 0 24 24'%3E%3Ccircle cx='11' cy='11' r='8'/%3E%3Cpath d='m21 21-4.35-4.35'/%3E%3C/svg%3E") no-repeat 12px center;
      transition: border-color .15s;
    }}
    #suche:focus {{ border-color: #2563eb; }}
  </style>
</head>
<body>
  {pw_overlay_html}
  <header>
    <h1>📚 Bibliothek Verlauf</h1>
    <div class="subtitle">{anz_gesamt} Medien insgesamt · {anz_aktiv} aktuell ausgeliehen · Stand {generiert}</div>
    <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-top:10px;">
      <input id="suche" type="search" placeholder="Titel oder Verfasser suchen…" oninput="applyVisibility()">
      <button id="aktiv-filter-btn" onclick="toggleFilter()">Nur aktuell ausgeliehene</button>
      {"".join(f'<button class="nutzer-filter-btn" data-n="{n.lower()}" onclick="toggleNutzer(this)">{n}</button>' for n in alle_nutzer)}
      <a href="statistik.html" style="font-size:0.82rem;color:#6b7280;text-decoration:none;margin-left:4px">📊 Statistik</a>
    </div>
  </header>
  {sektionen_html}
  <footer>Stadtbibliothek Halle · generiert von bibliothek_kalender_sync.py</footer>
  {pw_script_html}
  <script>
  function applyVisibility() {{
    const q = (document.getElementById('suche').value || '').toLowerCase().trim();
    const nurAktiv = sessionStorage.getItem('bib_filter') === '1';
    const nurNutzer = sessionStorage.getItem('bib_nutzer') || '';
    document.querySelectorAll('.karte').forEach(k => {{
      const titel = (k.querySelector('.titel')?.textContent || '').toLowerCase();
      const verf  = (k.querySelector('.verfasser')?.textContent || '').toLowerCase();
      const matchSuche  = !q || titel.includes(q) || verf.includes(q);
      const matchFilter = !nurAktiv || !k.classList.contains('karte-zurueck');
      const matchNutzer = !nurNutzer || k.dataset.nutzer === nurNutzer;
      k.style.display = (matchSuche && matchFilter && matchNutzer) ? '' : 'none';
    }});
    document.querySelectorAll('details').forEach(d => {{
      const hatSichtbare = [...d.querySelectorAll('.karte')].some(k => k.style.display !== 'none');
      d.style.display = hatSichtbare ? '' : 'none';
    }});
  }}
  function applyFilter(aktiv) {{
    sessionStorage.setItem('bib_filter', aktiv ? '1' : '0');
    const btn = document.getElementById('aktiv-filter-btn');
    btn.classList.toggle('aktiv', aktiv);
    btn.textContent = aktiv ? 'Alle anzeigen' : 'Nur aktuell ausgeliehene';
    applyVisibility();
  }}
  function toggleFilter() {{
    applyFilter(sessionStorage.getItem('bib_filter') !== '1');
  }}
  function toggleNutzer(btn) {{
    const n = btn.dataset.n;
    const aktiv = sessionStorage.getItem('bib_nutzer') === n;
    sessionStorage.setItem('bib_nutzer', aktiv ? '' : n);
    document.querySelectorAll('.nutzer-filter-btn').forEach(b => b.classList.remove('aktiv'));
    if (!aktiv) btn.classList.add('aktiv');
    applyVisibility();
  }}
  applyFilter(sessionStorage.getItem('bib_filter') === '1');
  const _n = sessionStorage.getItem('bib_nutzer');
  if (_n) {{ const b = document.querySelector(`.nutzer-filter-btn[data-n="${{_n}}"]`); if (b) b.classList.add('aktiv'); applyVisibility(); }}
  </script>
</body>
</html>"""

    ziel = VERLAUF_HTML_LOKAL if lokal else VERLAUF_HTML
    with open(ziel, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"   📄 {ziel} generiert ({anz_gesamt} Einträge, {len(gruppen)} Gruppen)")


# ─────────────────────────────────────────────
# SCHRITT 4: Statistik-Seite
# ─────────────────────────────────────────────

def generiere_statistik_html(root):
    """Generiert statistik.html mit Ausleihdaten aus verlauf.xml."""
    heute = datetime.date.today()
    generiert = heute.strftime("%d.%m.%Y")

    # Alle Ausleihen sammeln
    alle = []
    for elem in root.findall("medium"):
        titel        = _xml_text(elem, "titel")
        mediengruppe = _xml_text(elem, "mediengruppe")
        nutzer_medium = _xml_text(elem, "nutzer")
        ausleihen_elem = elem.find("ausleihen")
        if ausleihen_elem is None:
            continue
        for a in ausleihen_elem.findall("ausleihe"):
            nutzer  = _xml_text(a, "nutzer") or nutzer_medium
            seit    = _xml_text(a, "seit")
            zurueck = _xml_text(a, "zurueck")
            frist   = _xml_text(a, "frist")
            alle.append({"nutzer": nutzer, "titel": titel, "mediengruppe": mediengruppe,
                         "seit": seit, "zurueck": zurueck, "frist": frist})

    gesamt = len(alle)
    if gesamt == 0:
        return

    # Ausleihen pro Nutzer
    pro_nutzer = defaultdict(int)
    for a in alle:
        pro_nutzer[a["nutzer"]] += 1

    # Ausleihen pro Mediengruppe, aufgeschlüsselt nach Nutzer
    pro_gruppe_nutzer = defaultdict(lambda: defaultdict(int))
    for a in alle:
        pro_gruppe_nutzer[a["mediengruppe"]][a["nutzer"]] += 1
    pro_gruppe_nutzer = dict(sorted(pro_gruppe_nutzer.items(), key=lambda x: -sum(x[1].values())))

    # Ausleihen pro Monat (alle Nutzer zusammen)
    pro_monat = defaultdict(int)
    for a in alle:
        if a["seit"]:
            pro_monat[a["seit"][:7]] += 1
    pro_monat = dict(sorted(pro_monat.items()))

    # Durchschnittliche Ausleihdauer pro Nutzer (nur abgeschlossene)
    tage_pro_nutzer = defaultdict(list)
    for a in alle:
        if a["seit"] and a["zurueck"]:
            try:
                d = (datetime.date.fromisoformat(a["zurueck"]) - datetime.date.fromisoformat(a["seit"])).days
                if d >= 0:
                    tage_pro_nutzer[a["nutzer"]].append(d)
            except ValueError:
                pass

    # Medien mehrfach ausgeliehen
    ausleihen_pro_medium = defaultdict(list)
    for a in alle:
        ausleihen_pro_medium[a["titel"]].append(a)
    mehrfach = {t: v for t, v in ausleihen_pro_medium.items() if len(v) > 1}
    mehrfach = dict(sorted(mehrfach.items(), key=lambda x: -len(x[1])))

    def farbe(n):
        return NUTZER_FARBEN_STATISTIK.get(n, "#a78bfa")

    def balken_html(items, farbe_fn, max_val=None):
        if not items:
            return ""
        mv = max_val or max(items.values())
        rows = ""
        for label, val in items.items():
            pct = int(val / mv * 100) if mv else 0
            rows += (
                f'<div class="bar-row">'
                f'<span class="bar-label">{label}</span>'
                f'<div class="bar-track">'
                f'<div class="bar-fill" style="width:{pct}%;background:{farbe_fn(label)}"></div>'
                f'</div>'
                f'<span class="bar-val">{val}</span>'
                f'</div>'
            )
        return rows

    # Nutzer-Balken
    nutzer_balken = balken_html(dict(sorted(pro_nutzer.items(), key=lambda x: -x[1])), farbe)

    # Gruppen-Balken, pro Nutzer aufgeschlüsselt
    def gruppe_balken_html(pro_gruppe_nutzer):
        if not pro_gruppe_nutzer:
            return ""
        mv = max(sum(nv.values()) for nv in pro_gruppe_nutzer.values())
        alle_nutzer = sorted({n for nv in pro_gruppe_nutzer.values() for n in nv})
        rows = ""
        for gruppe, nv in pro_gruppe_nutzer.items():
            sub = "".join(
                f'<div class="bar-row" style="margin-bottom:4px">'
                f'<span class="bar-label" style="min-width:60px;font-size:0.75rem;color:{farbe(n)}">{n}</span>'
                f'<div class="bar-track">'
                f'<div class="bar-fill" style="width:{int(nv.get(n,0)/mv*100)}%;background:{farbe(n)}"></div>'
                f'</div>'
                f'<span class="bar-val">{nv.get(n,0)}</span>'
                f'</div>'
                for n in alle_nutzer if nv.get(n, 0) > 0
            )
            rows += (
                f'<div style="margin-bottom:14px">'
                f'<div style="font-size:0.82rem;font-weight:600;color:#374151;margin-bottom:4px">{gruppe}</div>'
                f'{sub}'
                f'</div>'
            )
        return rows
    gruppe_balken = gruppe_balken_html(pro_gruppe_nutzer)

    # Monats-Balken
    monat_farbe = lambda _: "#f59e0b"
    monat_balken = balken_html(pro_monat, monat_farbe)

    # Durchschnittsdauer-Tabelle
    dauer_rows = ""
    for n, tage in sorted(tage_pro_nutzer.items()):
        avg = round(sum(tage) / len(tage), 1)
        dauer_rows += f'<tr><td style="color:{farbe(n)};font-weight:600">{n}</td><td>{avg} Tage</td><td style="color:#9ca3af">(aus {len(tage)} Ausleihen)</td></tr>'

    # Mehrfach-Tabelle
    mehrfach_rows = ""
    for titel, eintraege in list(mehrfach.items())[:10]:
        badges = "".join(
            f'<span class="badge" style="background:{farbe(e["nutzer"])}20;color:{farbe(e["nutzer"])};border:1px solid {farbe(e["nutzer"])}40">'
            f'{e["nutzer"]} {e["seit"][:7] if e["seit"] else ""}</span>'
            for e in eintraege
        )
        mehrfach_rows += f'<tr><td class="mt-titel">{titel}</td><td>{len(eintraege)}×</td><td>{badges}</td></tr>'

    html = f"""<!DOCTYPE html>
<html lang="de">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Bibliothek Statistik</title>
  <style>
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #f0f2f5; color: #111827; padding: 24px 16px; }}
    header {{ text-align: center; margin-bottom: 28px; }}
    header h1 {{ font-size: 1.6rem; font-weight: 700; }}
    header .subtitle {{ font-size: 0.82rem; color: #6b7280; margin-top: 4px; }}
    .nav {{ text-align: center; margin-bottom: 24px; }}
    .nav a {{ font-size: 0.85rem; color: #3b82f6; text-decoration: none; }}
    .nav a:hover {{ text-decoration: underline; }}
    .grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 16px; max-width: 960px; margin: 0 auto; }}
    .card {{ background: #fff; border-radius: 14px; padding: 20px 22px; box-shadow: 0 1px 6px rgba(0,0,0,.07); }}
    .card h2 {{ font-size: 0.95rem; font-weight: 700; color: #374151; margin-bottom: 16px; }}
    .kpi-grid {{ display: flex; gap: 16px; flex-wrap: wrap; max-width: 960px; margin: 0 auto 16px; }}
    .kpi {{ background: #fff; border-radius: 14px; padding: 16px 20px; flex: 1; min-width: 120px; text-align: center; box-shadow: 0 1px 6px rgba(0,0,0,.07); }}
    .kpi .val {{ font-size: 2rem; font-weight: 800; }}
    .kpi .lbl {{ font-size: 0.75rem; color: #6b7280; margin-top: 2px; }}
    .bar-row {{ display: flex; align-items: center; gap: 10px; margin-bottom: 10px; }}
    .bar-label {{ font-size: 0.82rem; min-width: 110px; color: #374151; }}
    .bar-track {{ flex: 1; background: #f3f4f6; border-radius: 999px; height: 10px; overflow: hidden; }}
    .bar-fill {{ height: 100%; border-radius: 999px; transition: width .4s; }}
    .bar-val {{ font-size: 0.8rem; color: #6b7280; min-width: 24px; text-align: right; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 0.83rem; }}
    td {{ padding: 6px 4px; border-bottom: 1px solid #f3f4f6; vertical-align: top; }}
    .mt-titel {{ max-width: 180px; word-break: break-word; color: #111827; }}
    .badge {{ display: inline-block; padding: 2px 7px; border-radius: 999px; font-size: 0.72rem; margin: 2px 2px 0 0; white-space: nowrap; }}
    footer {{ text-align: center; font-size: 0.75rem; color: #9ca3af; margin-top: 28px; }}
  </style>
</head>
<body>
  <header>
    <h1>📊 Bibliothek Statistik</h1>
    <div class="subtitle">Stand {generiert} · {gesamt} Ausleihen gesamt</div>
  </header>
  <div class="nav"><a href="verlauf.html">← Zurück zum Verlauf</a></div>

  <div class="kpi-grid">
    {"".join(f'<div class="kpi"><div class="val" style="color:{farbe(n)}">{v}</div><div class="lbl">{n}</div></div>' for n, v in sorted(pro_nutzer.items(), key=lambda x: -x[1]))}
    <div class="kpi"><div class="val" style="color:#6366f1">{len(pro_gruppe_nutzer)}</div><div class="lbl">Medientypen</div></div>
    <div class="kpi"><div class="val" style="color:#f59e0b">{len(pro_monat)}</div><div class="lbl">Monate aktiv</div></div>
  </div>

  <div class="grid">
    <div class="card">
      <h2>Ausleihen pro Nutzer</h2>
      {nutzer_balken}
    </div>
    <div class="card">
      <h2>Mediengruppen</h2>
      {gruppe_balken}
    </div>
    <div class="card" style="grid-column: 1 / -1">
      <h2>Ausleihen pro Monat</h2>
      {monat_balken}
    </div>
    {'<div class="card"><h2>Ø Ausleihdauer</h2><table>' + dauer_rows + '</table></div>' if dauer_rows else ''}
    {'<div class="card" style="grid-column: 1 / -1"><h2>Mehrfach ausgeliehen</h2><table>' + mehrfach_rows + '</table></div>' if mehrfach_rows else ''}
  </div>

  <footer>Generiert am {generiert}</footer>
</body>
</html>"""

    with open(VERLAUF_STATISTIK, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"   📊 {VERLAUF_STATISTIK} generiert")


# ─────────────────────────────────────────────
# SCHRITT 5: Hauptprogramm
# ─────────────────────────────────────────────

def main():
    print("\n" + "═" * 55)
    print("  📚 Bibliothek Halle → Google Calendar Sync & Verlauf")
    print("═" * 55 + "\n")

    # 1. Medien aller Benutzer abrufen
    alle_medien = []
    abruf_fehlgeschlagen = False
    for benutzer in BENUTZER:
        print(f"─── Ausgeliehene Medien: {benutzer['name']} ───")
        try:
            medien = hole_ausgeliehene_medien(
                benutzer["name"], benutzer["ausweis"], benutzer["passwort"]
            )
            alle_medien.extend(medien)
        except Exception as e:
            print(f"\n❌ Fehler beim Abruf für {benutzer['name']}:\n   {e}")
            abruf_fehlgeschlagen = True

    if abruf_fehlgeschlagen:
        print("\n⚠️  Sync abgebrochen – Bibliotheksportal nicht vollständig erreichbar.")
        print("   Kalender und Verlauf bleiben unverändert.")
        return

    if not alle_medien:
        print("ℹ️  Aktuell keine Medien ausgeliehen.")

    # 2. Nach Datum gruppieren
    medien_nach_datum = {}
    for medium in alle_medien:
        medien_nach_datum.setdefault(medium["frist"].isoformat(), []).append(medium)

    # 3. Google Calendar verbinden
    print("\n" + "═" * 55)
    print("─── Google Calendar ───")
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
    print("\n" + "═" * 55)
    print("─── Verlauf aktualisieren ───")
    verlauf_root = aktualisiere_verlauf(alle_medien)
    generiere_html(verlauf_root)
    generiere_statistik_html(verlauf_root)
    if not os.environ.get("GITHUB_ACTIONS"):
        generiere_html(verlauf_root, lokal=True)

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
