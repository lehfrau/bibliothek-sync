#!/usr/bin/env python3
"""
Stadtbibliothek Halle → Google Calendar Sync
=============================================
Liest ausgeliehene Medien aus dem Bibliotheksportal (katalog.halle.de)
und erstellt/aktualisiert Kalendereinträge in Google Calendar.

Setup-Anleitung: siehe README-Abschnitt am Ende dieser Datei.
"""

import os
import re
import json
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

BIBLIOTHEK_AUSWEISNUMMER = os.environ.get("BIBL_AUSWEIS", "")
BIBLIOTHEK_PASSWORT      = os.environ.get("BIBL_PASSWORT", "")

GOOGLE_CREDENTIALS_FILE  = "credentials.json"
GOOGLE_TOKEN_FILE        = "token.json"

# ID des Ziel-Kalenders (oder "primary" für Hauptkalender)
# Eigenen Kalender empfohlen, z.B. "Bibliothek Fristen"
KALENDER_ID = "sc2ufkcp6satu4qnrq4mehj01k@group.calendar.google.com"

# Wie viele Tage vor Ablauf soll der Erinnerungsalarm erscheinen?
ERINNERUNG_TAGE_VORHER = 3

# Kennzeichnung für Kalendereinträge (damit das Script nur eigene Einträge anfasst)
EVENT_KENNZEICHEN = "📚 Bibliothek Halle"

# ─────────────────────────────────────────────
# KONSTANTEN
# ─────────────────────────────────────────────

LOGIN_URL   = "https://katalog.halle.de/Mein-Konto"
KONTO_URL   = "https://katalog.halle.de/Mein-Konto"
SCOPES      = ["https://www.googleapis.com/auth/calendar"]


# ─────────────────────────────────────────────
# SCHRITT 1: Bibliothekskonto scrapen
# ─────────────────────────────────────────────

def hole_ausgeliehene_medien():
    """
    Loggt sich ins Bibliotheksportal ein und liest alle ausgeliehenen Medien
    mit Titel und Rückgabedatum aus.

    Rückgabe: Liste von Dicts mit keys: titel, frist (datetime.date), medium_id
    """
    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )
    })

    # Seite laden um VIEWSTATE etc. zu holen
    print("📡 Verbinde mit Bibliotheksportal...")
    response = session.get(LOGIN_URL)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")

    # ASP.NET-Formularfelder auslesen
    def get_hidden(name):
        tag = soup.find("input", {"name": name})
        return tag["value"] if tag else ""

    viewstate          = get_hidden("__VIEWSTATE")
    viewstate_gen      = get_hidden("__VIEWSTATEGENERATOR")
    event_validation   = get_hidden("__EVENTVALIDATION")
    viewstate_enc      = get_hidden("__VIEWSTATEENCRYPTED")

    # Login-Formular absenden
    # Feldnamen aus Debug-Ausgabe ermittelt (katalog.halle.de)
    login_field_user = "dnn$ctr375$Login$Login_COP$txtUsername"
    login_field_pass = "dnn$ctr375$Login$Login_COP$txtPassword"

    post_data = {
        "__EVENTTARGET":        "",
        "__EVENTARGUMENT":      "",
        "__VIEWSTATE":          viewstate,
        "__VIEWSTATEGENERATOR": viewstate_gen,
        "__EVENTVALIDATION":    event_validation,
        "__VIEWSTATEENCRYPTED": viewstate_enc,
        login_field_user:       BIBLIOTHEK_AUSWEISNUMMER,
        login_field_pass:       BIBLIOTHEK_PASSWORT,
        "dnn$ctr375$Login$Login_COP$cmdLogin": "Anmelden",
    }

    print("🔐 Melde an...")
    response = session.post(LOGIN_URL, data=post_data)
    response.raise_for_status()

    # Prüfen ob Login erfolgreich (Name erscheint im Header)
    if "abmelden" not in response.text.lower():
        raise RuntimeError(
            "❌ Login fehlgeschlagen. Bitte Ausweisnummer und Passwort prüfen.\n"
            "   Hinweis: Das Portal könnte eine andere Login-Struktur haben –\n"
            "   dann run_debug_login() aufrufen für Diagnose."
        )
    print("✅ Login erfolgreich.")

    return parse_ausleih_seite(response.text)


def parse_ausleih_seite(html):
    """
    Parst die HTML-Seite und extrahiert ausgeliehene Medien.
    Erwartet Tabelle mit ID 'grdViewLoans'.
    """
    soup = BeautifulSoup(html, "html.parser")
    medien = []

    # Ausleih-Tabelle finden
    tabelle = soup.find("table", {"id": lambda x: x and "grdViewLoans" in x})

    if not tabelle:
        print("ℹ️  Keine ausgeliehenen Medien gefunden (Tabelle leer oder nicht vorhanden).")
        return medien

    zeilen = tabelle.find_all("tr")

    for zeile in zeilen[1:]:  # erste Zeile = Header überspringen
        zellen = zeile.find_all("td")
        if len(zellen) < 4:
            continue

        # Titel aus Link extrahieren
        titel_link = zeile.find("a", href=lambda h: h and "Mediensuche" in h)
        if not titel_link:
            continue
        titel = titel_link.get_text(strip=True)

        # Medium-ID aus href
        href = titel_link.get("href", "")
        id_match = re.search(r"id=(\d+)", href)
        medium_id = id_match.group(1) if id_match else titel

        # Datum suchen: Format TT.MM.JJJJ
        # Die Frist steht in der Zelle, die ein Datum im deutschen Format enthält
        datum = None
        for zelle in zellen:
            text = zelle.get_text(strip=True)
            datum_match = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", text)
            if datum_match:
                tag, monat, jahr = datum_match.groups()
                kandidat = datetime.date(int(jahr), int(monat), int(tag))
                # Nehme das späteste Datum (= Rückgabefrist, nicht Ausleihdatum)
                if datum is None or kandidat > datum:
                    datum = kandidat

        if titel and datum:
            medien.append({
                "titel":     titel,
                "frist":     datum,
                "medium_id": medium_id,
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
    Holt alle bestehenden Bibliotheks-Kalendereinträge (erkennbar am Kennzeichen).
    Rückgabe: Dict {medium_id: event_dict}
    """
    events_by_id = {}
    page_token = None

    while True:
        result = service.events().list(
            calendarId=KALENDER_ID,
            q=EVENT_KENNZEICHEN,
            singleEvents=True,
            maxResults=250,
            pageToken=page_token,
        ).execute()

        for event in result.get("items", []):
            # Medium-ID aus extended properties lesen
            props = event.get("extendedProperties", {}).get("private", {})
            mid = props.get("medium_id")
            if mid:
                events_by_id[mid] = event

        page_token = result.get("nextPageToken")
        if not page_token:
            break

    return events_by_id


def erstelle_oder_aktualisiere_event(service, medium, bestehende_events):
    """
    Erstellt einen neuen Kalendereintrag oder verschiebt einen bestehenden,
    wenn sich das Datum geändert hat (= Verlängerung).
    """
    medium_id   = str(medium["medium_id"])
    titel       = medium["titel"]
    frist       = medium["frist"]
    summary     = f"📚 {titel}"

    # Datum als ISO-String
    datum_str = frist.isoformat()

    # Erinnerung: X Tage vor Ablauf in Minuten umrechnen
    erinnerung_minuten = ERINNERUNG_TAGE_VORHER * 24 * 60

    neues_event = {
        "summary": summary,
        "description": (
            f"Rückgabefrist: {frist.strftime('%d.%m.%Y')}\n"
            f"Bibliothek Halle – Katalog: https://katalog.halle.de/Mediensuche?id={medium_id}"
        ),
        "start": {"date": datum_str},
        "end":   {"date": datum_str},
        "reminders": {
            "useDefault": False,
            "overrides": [
                {"method": "popup",  "minutes": erinnerung_minuten},
                {"method": "email",  "minutes": erinnerung_minuten},
            ],
        },
        "extendedProperties": {
            "private": {
                "medium_id":        medium_id,
                "bibliothek_sync":  "true",
            }
        },
        "colorId": "9",  # Blaubeere – gut erkennbar
    }

    if medium_id in bestehende_events:
        altes_event = bestehende_events[medium_id]
        altes_datum = altes_event.get("start", {}).get("date", "")

        if altes_datum == datum_str:
            print(f"   ⏭️  Unverändert: {titel} ({frist.strftime('%d.%m.%Y')})")
            return

        # Datum hat sich geändert → Event aktualisieren (Verlängerung!)
        event_id = altes_event["id"]
        service.events().update(
            calendarId=KALENDER_ID,
            eventId=event_id,
            body=neues_event,
        ).execute()
        print(
            f"   🔄 Verlängert: {titel}\n"
            f"      {altes_datum} → {datum_str}"
        )
    else:
        # Neues Event anlegen
        service.events().insert(
            calendarId=KALENDER_ID,
            body=neues_event,
        ).execute()
        print(f"   ✅ Neu: {titel} → {frist.strftime('%d.%m.%Y')}")


def loesche_veraltete_events(service, aktuelle_medium_ids, bestehende_events):
    """
    Löscht Kalendereinträge für Medien, die nicht mehr ausgeliehen sind
    (d.h. zurückgegeben wurden).
    """
    for medium_id, event in bestehende_events.items():
        if medium_id not in aktuelle_medium_ids:
            titel = event.get("summary", medium_id)
            service.events().delete(
                calendarId=KALENDER_ID,
                eventId=event["id"],
            ).execute()
            print(f"   🗑️  Zurückgegeben (Event gelöscht): {titel}")


# ─────────────────────────────────────────────
# SCHRITT 3: Hauptprogramm
# ─────────────────────────────────────────────

def main():
    print("\n" + "═" * 55)
    print("  📚 Bibliothek Halle → Google Calendar Sync")
    print("═" * 55 + "\n")

    # 1. Medien aus Bibliotheksportal holen
    print("─── Ausgeliehene Medien ───")
    try:
        medien = hole_ausgeliehene_medien()
    except Exception as e:
        print(f"\n❌ Fehler beim Abruf des Bibliothekskontos:\n   {e}")
        return

    if not medien:
        print("ℹ️  Aktuell keine Medien ausgeliehen.")

    # 2. Google Calendar verbinden
    print("\n─── Google Calendar ───")
    try:
        service = google_calendar_service()
    except Exception as e:
        print(f"\n❌ Fehler bei Google-Authentifizierung:\n   {e}")
        return

    # 3. Bestehende Bibliotheks-Events laden
    bestehende_events = hole_bestehende_biblio_events(service)
    print(f"   Gefunden: {len(bestehende_events)} bestehende Einträge im Kalender.\n")

    # 4. Events erstellen/aktualisieren
    print("─── Kalender aktualisieren ───")
    aktuelle_ids = set()
    for medium in medien:
        mid = str(medium["medium_id"])
        aktuelle_ids.add(mid)
        erstelle_oder_aktualisiere_event(service, medium, bestehende_events)

    # 5. Zurückgegebene Medien aus Kalender entfernen
    loesche_veraltete_events(service, aktuelle_ids, bestehende_events)

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
        name = inp.get("name", "")
        typ  = inp.get("type", "text")
        if typ not in ("hidden",):
            print(f"  [{typ}] name='{name}'")

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
# Oben im Script BIBLIOTHEK_AUSWEISNUMMER und BIBLIOTHEK_PASSWORT setzen.
# Alternativ: Umgebungsvariablen BIBL_AUSWEIS und BIBL_PASSWORT setzen
# und die Zeilen oben durch diese ersetzen:
#
#   BIBLIOTHEK_AUSWEISNUMMER = os.environ["BIBL_AUSWEIS"]
#   BIBLIOTHEK_PASSWORT      = os.environ["BIBL_PASSWORT"]
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
# (sieht aus wie: abc123@group.calendar.google.com)
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
