# RWK Datenbank-Import (DISAG OpticScore - Server)

Dieses Python-Script erzeugt eine neue Rundenwettkampf-Datenbank für das Tool
**DISAG OpticScore - Server, Version 1.60.9.1** und befüllt sie automatisch aus
CSV-Exporten (Vereine, Mannschaften, Schützen, Wettkämpfe).

Es ersetzt das manuelle Anlegen von Vereinen, Mannschaften, Schützen und
Wettkämpfen in der Access-Datenbank durch einen einmaligen, wiederholbaren
Import-Lauf.

## Was das Script macht

1. Kopiert eine leere Access-Vorlage (`RWK_leer/RWK_leer.mdb`) als neue
   Zieldatenbank (falls diese noch nicht existiert).
2. Liest die vier CSV-Dateien (Verein, Mannschaften, Schützen, Wettkämpfe) ein.
3. Befüllt in der Zieldatenbank u. a. folgende Tabellen:
   - `Clubs`
   - `Teams` (inkl. konfigurierbarem Mannschaftsnamen-Format)
   - `Shooters` / `TeamsShooters`
   - `Leaguecompetitions_Competitions`
   - `Leaguecompetitions_Shooters`

## Voraussetzungen

- Windows mit installiertem **Microsoft Access Driver (\*.mdb, \*.accdb)**
  (64-Bit ODBC-Treiber, Teil des Access Database Engine Redistributable).
- Python 3.10 oder neuer.
- Virtuelle Umgebung mit dem Paket `pyodbc`, z. B.:

  ```powershell
  python -m venv .venv
  .\.venv\Scripts\pip.exe install pyodbc
  ```

- Eine leere Vorlagen-Datenbank (`RWK_leer/RWK_leer.mdb`), z. B. exportiert aus
  DISAG OpticScore - Server.
- Die vier CSV-Dateien (Verein, Mannschaften, Schützen, Wettkämpfe) im in
  `config.ini` konfigurierten Ordner, mit Semikolon als Trennzeichen und
  UTF-8-Kodierung (BOM).

## Konfiguration

Alle Pfade, Dateinamen, Spaltennamen und das Format der Mannschaftsnamen
werden in `config.ini` festgelegt (siehe Kommentare in der Datei). Insbesondere:

- `[RWK] Verein` bestimmt per Regex den eigenen Verein (für Log-Ausgaben).
- `[Pfade]` bestimmt CSV-Ordner, Vorlage und Ziel-Datenbank.
- `[CSV_Dateien]` legt Dateinamen und erwartete Spalten je CSV fest.
- `[Teams] teamsname_format` / `teamsnameshort_format` steuern die
  Mannschaftsnamen, z. B. mit den Platzhaltern `{klassenname}`,
  `{klasse_ohne_zahl}`, `{nummer}`, `{vereinsid}`, `{mannschaftsid}`,
  `{vereinsname}`, `{vereinsort}`.

## Verwendung

```powershell
.\.venv\Scripts\python.exe .\import_rwk.py [config.ini]
```

Um die Zieldatenbank komplett neu zu erzeugen (statt nur die Inhalte zu
ersetzen), vorher die vorhandene Ziel-Datei löschen:

```powershell
Remove-Item .\RWK_2026\RWK_2026.mdb, .\RWK_2026\RWK_2026.sdf -ErrorAction SilentlyContinue
.\.venv\Scripts\python.exe .\import_rwk.py
```

Das Script gibt am Ende Warnungen aus, z. B. bei gekürzten Feldern oder
mehrdeutigen Mannschafts-Zuordnungen.

## Lizenz und Haftungsausschluss

Dieses Projekt darf frei als Open Source verwendet, verändert und
weiterverbreitet werden. Es besteht **kein Anspruch auf Richtigkeit,
Vollständigkeit oder Eignung für einen bestimmten Zweck** – die Nutzung
erfolgt auf eigenes Risiko und ohne jegliche Gewährleistung.

## Hinweis

Dieses Script wurde unter Verwendung von GitHub Copilot erstellt.
