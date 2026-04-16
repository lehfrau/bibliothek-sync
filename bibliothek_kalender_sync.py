#!/usr/bin/env python3
"""
Stadtbibliothek Halle → Google Calendar Sync
=============================================
Liest ausgeliehene Medien aller Benutzer aus dem Bibliotheksportal (katalog.halle.de)
und erstellt/aktualisiert Kalendereinträge in Google Calendar.
Medien mit gleichem Rückgabedatum werden in einem gemeinsamen Eintrag zusammengefasst.

Setup-Anleitung: siehe README-Abschnitt am Ende dieser Datei.
"""

import os
import re
import datetime
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
    Rückgabe: Liste von Dicts mit keys: titel, frist (datetime.date), medium_id, name
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

        if titel and datum:
            medien.append({
                "titel":     titel,
                "frist":     datum,
                "medium_id": medium_id,
                "name":      name,
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
# SCHRITT 3: Hauptprogramm
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

    # 2. Medien nach Rückgabedatum gruppieren
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