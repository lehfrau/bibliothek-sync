# Bibliothek Halle → Google Calendar: Einrichtungsanleitung

**Ziel:** Ausleihfristen von katalog.halle.de werden automatisch in Google Calendar eingetragen – mit Erinnerung 3 Tage vor Ablauf. Verlängerungen werden automatisch aktualisiert, Rückgaben automatisch gelöscht.

**Läuft automatisch:** jeden Montag und Donnerstag

---

## Was du brauchst

- Windows-PC
- Bibliotheksausweisnummer + Passwort
- Google-Konto
- GitHub-Konto (kostenlos)

---

## Phase 1: Python einrichten

### 1.1 Python installieren

1. Gehe auf **https://www.python.org/downloads/**
2. Klick auf den großen Download-Button
3. Öffne die heruntergeladene `.exe`
4. ⚠️ **Wichtig:** Häkchen bei „Add Python to PATH" setzen!
5. „Install Now" klicken

### 1.2 Installation prüfen

`Windows + R` → `cmd` → Enter → eintippen:

```
python --version
```

→ Es erscheint z.B. `Python 3.12.3` ✅

### 1.3 Pakete installieren

In der Eingabeaufforderung:

```
pip install requests beautifulsoup4 google-auth google-auth-oauthlib google-api-python-client
```

→ Endet mit „Successfully installed ..." ✅

---

## Phase 2: Script einrichten

### 2.1 Ordner anlegen

Neuen Ordner erstellen, z.B. auf dem Desktop: `Bibliothek-Sync`

### 2.2 Script herunterladen

Die Datei `bibliothek_kalender_sync.py` in den Ordner legen.

### 2.3 Zugangsdaten eintragen

Datei mit Texteditor öffnen (Rechtsklick → „Mit Editor öffnen").

Diese zwei Zeilen suchen und anpassen:

```python
BIBLIOTHEK_AUSWEISNUMMER = os.environ.get("BIBL_AUSWEIS", "")
BIBLIOTHEK_PASSWORT      = os.environ.get("BIBL_PASSWORT", "")
```

Für lokale Tests auf dem eigenen PC: In der Eingabeaufforderung einmalig setzen:

```
set BIBL_AUSWEIS=DEINE_AUSWEISNUMMER
set BIBL_PASSWORT=DEIN_PASSWORT
```

### 2.4 Login testen

In der Eingabeaufforderung in den Ordner navigieren:

```
cd C:\Users\DEINNAME\Desktop\Bibliothek-Sync
python bibliothek_kalender_sync.py --debug
```

→ Es erscheint eine Liste von Formularfeldern. Darunter muss auftauchen:

```
[text]     name='dnn$ctr375$Login$Login_COP$txtUsername'
[password] name='dnn$ctr375$Login$Login_COP$txtPassword'
```

✅ Wenn ja: Feldnamen stimmen, Login funktioniert.

---

## Phase 3: Google Calendar API einrichten

### 3.1 Google Cloud Console öffnen

Gehe auf **https://console.cloud.google.com/**  
Mit dem Google-Konto anmelden, in dessen Kalender die Einträge sollen.

### 3.2 Neues Projekt erstellen

Oben links Dropdown → „Neues Projekt" → Name: `Bibliothek Sync` → „Erstellen"

### 3.3 Google Calendar API aktivieren

Linkes Menü → „APIs & Dienste" → „Bibliothek" → `Google Calendar API` suchen → „Aktivieren"

### 3.4 OAuth einrichten

Linkes Menü → „APIs & Dienste" → „OAuth-Zustimmungsbildschirm" (oder: „Google Auth Platform" → „Get started")

Ausfüllen:
- App-Name: `Bibliothek Sync`
- Support-E-Mail: eigene E-Mail-Adresse
- Audience: `External`
- Nutzungsbedingungen akzeptieren
- Durchklicken bis fertig

### 3.5 OAuth-Client erstellen

„Clients" → „Create OAuth client"
- Application type: `Desktop app`
- Name: `Bibliothek Sync`
- „Create" → „OK"

### 3.6 Testnutzer eintragen

„Audience" → „Test users" → „Add users" → eigene Google-E-Mail-Adresse eintragen → „Save"

⚠️ Ohne diesen Schritt erscheint später „Error 403: access_denied"!

### 3.7 credentials.json herunterladen

„Clients" → auf „Bibliothek Sync" klicken → „Download JSON"

Datei umbenennen in `credentials.json` → in den `Bibliothek-Sync` Ordner legen.

### 3.8 Ersten Sync-Test durchführen

```
python bibliothek_kalender_sync.py
```

→ Ein Browser-Fenster öffnet sich → Google-Konto auswählen  
→ Warnung „App nicht überprüft" → „Erweitert" → „Weiter zu Bibliothek Sync"  
→ „Zulassen"

Im Terminal erscheint:
```
✅ Login erfolgreich.
✅ Sync abgeschlossen!
```

Im Ordner liegt jetzt automatisch eine `token.json` → nicht löschen, wird gebraucht!

---
## Kalender ID
1. Gehe auf https://calendar.google.com
2. Links in der Seitenleiste bei „Routinen" auf die drei Punkte klicken → „Einstellungen und Freigabe"
2. Ganz unten auf der Seite: Abschnitt „Kalender-ID" — sieht aus wie abc123xyz@group.calendar.google.com oder einfach deine E-Mail-Adresse wenn es ein persönlicher Kalender ist

Diese ID dann im Script eintragen:
`pythonKALENDER_ID = "abc123xyz@group.calendar.google.com"`

Statt dem bisherigen:
`pythonKALENDER_ID = "primary"`

## Phase 4: GitHub Actions (automatischer Betrieb)

### 4.1 Neues Repository erstellen

Gehe auf **https://github.com/new**
- Name: `bibliothek-sync`
- **Private** auswählen ⚠️
- „Create repository"

### 4.2 Script hochladen

Auf der Repository-Seite: „uploading an existing file" → `bibliothek_kalender_sync.py` hochladen → „Commit changes"

⚠️ `credentials.json` und `token.json` **nicht** ins Repo hochladen — diese kommen als Secrets rein (nächster Schritt).

### 4.3 Workflow-Datei erstellen

URL aufrufen: `https://github.com/DEIN-USERNAME/bibliothek-sync/new/main`

Im Namensfeld eintragen:
```
.github/workflows/sync.yml
```

Inhalt:

```yaml
name: Bibliothek Kalender Sync

on:
  schedule:
    - cron: "0 6 * * 1,4"
  workflow_dispatch:

jobs:
  sync:
    runs-on: ubuntu-latest

    steps:
      - name: Repository auschecken
        uses: actions/checkout@v4.2.2

      - name: Python einrichten
        uses: actions/setup-python@v5.6.0
        with:
          python-version: "3.11"

      - name: Abhängigkeiten installieren
        run: |
          pip install requests beautifulsoup4 \
            google-auth google-auth-oauthlib google-api-python-client

      - name: Google Token aus Secret wiederherstellen
        run: |
          echo '${{ secrets.GOOGLE_TOKEN_JSON }}' > token.json

      - name: Google Credentials aus Secret wiederherstellen
        run: |
          echo '${{ secrets.GOOGLE_CREDENTIALS_JSON }}' > credentials.json

      - name: Sync ausführen
        env:
          BIBL_AUSWEIS:  ${{ secrets.BIBL_AUSWEIS }}
          BIBL_PASSWORT: ${{ secrets.BIBL_PASSWORT }}
        run: python bibliothek_kalender_sync.py
```

„Commit changes" klicken.

### 4.4 Secrets hinterlegen

Repository → „Settings" → „Secrets and variables" → „Actions" → „New repository secret"

Folgende 4 Secrets anlegen:

| Name | Inhalt |
|------|--------|
| `BIBL_AUSWEIS` | Bibliotheksausweisnummer |
| `BIBL_PASSWORT` | Bibliothekspasswort |
| `GOOGLE_CREDENTIALS_JSON` | Kompletter Inhalt der `credentials.json` (Strg+A, Strg+C) |
| `GOOGLE_TOKEN_JSON` | Kompletter Inhalt der `token.json` (Strg+A, Strg+C) |

### 4.5 Ersten automatischen Lauf testen

Repository → „Actions" → „Bibliothek Kalender Sync" → „Run workflow" → „Run workflow"

→ Nach ca. 1 Minute: grüner Haken ✅ = alles funktioniert!

---

## Für einen zweiten Account (z.B. Benny)

Alles ab **Phase 3** wiederholen — mit Bennys Bibliothekszugangsdaten und seinem Google-Konto.

Optionen:
- **Eigenes Repo** für Benny (`bibliothek-sync-benny`) mit eigenen Secrets
- **Oder:** Bennys Daten als zusätzliche Secrets im selben Repo, Script leicht anpassen

Empfehlung: eigenes Repo, dann sind die Accounts sauber getrennt.

---

## Wartung & Troubleshooting

| Problem | Lösung |
|---------|--------|
| Login schlägt fehl | `--debug` ausführen, Feldnamen prüfen |
| Error 403: access_denied | Testnutzer in Google Auth Platform eintragen |
| token.json abgelaufen (alle ~6 Monate) | Lokal neu einloggen, `token.json` als Secret aktualisieren |
| Keine Medien gefunden obwohl ausgeliehen | Portal-HTML hat sich geändert, Script anpassen |
| GitHub Actions Warnung zu Node.js | Workflow-Versionen auf neueste aktualisieren |
