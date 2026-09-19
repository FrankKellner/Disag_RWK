# RWK Datenbank-Import (DISAG OpticScore - Server)

Dieses Python-Script erzeugt eine neue Rundenwettkampf-Datenbank für das Tool
**DISAG OpticScore - Server, Version 1.70.1** und befüllt sie automatisch aus
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
4. Exportiert die eigenen Heimwettkämpfe (mit aufgelösten Mannschaftsnamen) in
   eine separate CSV-Datei (Ordner/Dateiname konfigurierbar).

Hat ein Verein mehrere Mannschaften in derselben Klasse (z. B. bei
Vereins-Derbys), fragt das Script beim ersten betroffenen Wettkampf (mit
Datum und Klasse) interaktiv auf der Konsole nach, welche Mannschaftsnummer
gemeint ist. Bei weiteren Wettkämpfen mit derselben Klasse/demselben Verein
wird automatisch abwechselnd die jeweils andere Mannschaft verwendet, ohne
erneut nachzufragen.

## Voraussetzungen

- Windows mit installiertem **Microsoft Access Driver (\*.mdb, \*.accdb)**
  (64-Bit ODBC-Treiber, Teil des Access Database Engine Redistributable).
- Python 3.10 oder neuer, im PATH verfügbar (`python` oder `py`).
- Virtuelle Umgebung mit dem Paket `pyodbc` (siehe `requirements.txt`). Wird
  beim Start über `Import_starten.bat` automatisch angelegt/ergänzt, falls
  nicht vorhanden. Manuell geht das so:

  ```powershell
  python -m venv .venv
  .\.venv\Scripts\pip.exe install -r requirements.txt
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
- `[Export] heimwettkaempfe_ordner` / `heimwettkaempfe_datei` legen fest, wohin
  die CSV-Datei mit den eigenen Heimwettkämpfen (inkl. aufgelöster
  Mannschaftsnamen) exportiert wird.
- `[Logging] debug_level` steuert die Ausführlichkeit der Konsolenausgabe:
  `error` (Standard, keine `WARNUNG:`-Meldungen wie z. B. gekürzte Felder),
  `warnung` oder `debug` (zeigen zusätzlich `WARNUNG:`-Meldungen an).
- `[Klassen_Uebersetzung]` ist eine optionale Übersetzungstabelle für den
  Klassennamen **ohne** die abschließende Mannschaftsnummer (z. B.
  `Alters AUF = Alteraufgelegt`; die Nummer wird beim Anzeigen automatisch
  wieder angehängt). Das Script legt diesen Abschnitt beim ersten Lauf
  automatisch an (Identitätsabbildung) und ergänzt ihn bei neuen
  Klassennamen-Stämmen aus `Wettkaempfe_<Jahr>.csv` – die Übersetzungen können
  danach frei angepasst werden. Sie wirken sich auf die Mannschaftsnamen und
  den Export der Heimwettkämpfe aus.


## Verwendung

### Per Doppelklick (Windows)

`Import_starten.bat` doppelklicken. Beim ersten Start werden die virtuelle
Umgebung `.venv` und das Paket `pyodbc` automatisch angelegt/installiert,
falls sie noch fehlen. Danach löscht das Script eine eventuell vorhandene
Zieldatenbank (`.mdb`/`.sdf`) und führt den Import neu aus. Am Ende bleibt
das Konsolenfenster offen, damit Warnungen/Fehler gelesen werden können.

### Über die Kommandozeile

```powershell
.\.venv\Scripts\python.exe .\import_rwk.py [--neu] [config.ini]
```

Mit `--neu` wird die Zieldatenbank vor dem Import gelöscht und aus der
Vorlage neu angelegt (statt nur die Inhalte in der bestehenden Datei zu
ersetzen), z. B.:

```powershell
.\.venv\Scripts\python.exe .\import_rwk.py --neu
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
