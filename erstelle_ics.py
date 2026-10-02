"""Erzeugt pro eigener Mannschaft eine ICS-Kalenderdatei mit allen Heim- und
Auswaertswettkaempfen der Saison (aus Wettkaempfe_<Jahr>.csv).

Aufruf:  .venv\\Scripts\\python.exe erstelle_ics.py [config.ini]
"""
from __future__ import annotations

import re
import sys
import uuid
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

import import_rwk as rwk

BASE_DIR = rwk.BASE_DIR

VTIMEZONE_EUROPE_BERLIN = """BEGIN:VTIMEZONE
TZID:Europe/Berlin
X-LIC-LOCATION:Europe/Berlin
BEGIN:DAYLIGHT
TZOFFSETFROM:+0100
TZOFFSETTO:+0200
TZNAME:CEST
DTSTART:19700329T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:+0200
TZOFFSETTO:+0100
TZNAME:CET
DTSTART:19701025T030000
RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU
END:STANDARD
END:VTIMEZONE"""
ICS_PRODID = "PRODID:-//Disag_RWK//erstelle_ics.py//DE"


def build_teams_by_key(
    mannschaft_rows: list[dict],
    config,
    club_info: dict[str, dict[str, str]],
    translation_map: dict[str, str],
) -> dict[tuple[str, str], list[tuple[int, str, int]]]:
    """Gruppiert Mannschaften nach (Klassenname, Vereinsid), analog zu import_teams()
    in import_rwk.py, aber ohne Datenbankzugriff."""
    name_format = config["Teams"]["teamsname_format"]
    teams_by_key: dict[tuple[str, str], list[tuple[int, str, int]]] = defaultdict(list)
    for row in mannschaft_rows:
        vereinsid = row["Vereinsid"].strip()
        klassenname = row["Klassenname"].strip()
        klassenname_anzeige = rwk.translate_klassenname(translation_map, klassenname)
        nummer = int(row["Mannschaftsnummer"])
        idteams = int(row["Mannschaftsid"])
        club = club_info.get(vereinsid, {})
        format_vars = {
            "klassenname": klassenname_anzeige,
            "klasse_ohne_zahl": rwk.strip_trailing_number(klassenname_anzeige),
            "nummer": nummer,
            "vereinsid": vereinsid,
            "mannschaftsid": row["Mannschaftsid"].strip(),
            "vereinsname": club.get("name", ""),
            "vereinsort": club.get("ort", ""),
        }
        teamsname = name_format.format(**format_vars)
        teams_by_key[(klassenname, vereinsid)].append((nummer, teamsname, idteams))
    for candidates in teams_by_key.values():
        candidates.sort(key=lambda c: c[0])
    return teams_by_key


def parse_uhrzeit(value: str) -> tuple[int, int]:
    stunde, minute = value.strip().split(":")
    return int(stunde), int(minute)


def bestimme_disziplin(klassenname: str) -> str:
    """Leitet das Disziplin-Kuerzel (LP/KK/LGA/AN/LG) aus dem Klassennamen-Stamm ab."""
    stamm, _ = rwk.split_klassenname(klassenname)
    if stamm == "Offen LP":
        return "LP"
    if stamm == "Offen KK":
        return "KK"
    if stamm.endswith("AUF"):
        return "LGA"
    if stamm == "Senioren AN":
        return "AN"
    return "LG"


def sichere_dateiname(name: str) -> str:
    name = re.sub(r"[^\w\-. ]+", "_", name)
    return name.strip() or "Mannschaft"


def ics_escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")


def format_dt(dt: datetime) -> str:
    return dt.strftime("%Y%m%dT%H%M%S")


def sammle_termine(
    wettkampf_rows: list[dict],
    own_id: int,
    saison: str,
    teams_by_key: dict[tuple[str, str], list[tuple[int, str, int]]],
    club_info: dict[str, dict[str, str]],
    translation_map: dict[str, str],
    start_uhrzeit: tuple[int, int],
    end_uhrzeit: tuple[int, int],
) -> tuple[dict[int, list[dict]], dict[int, str]]:
    resolver = rwk.TeamResolver(teams_by_key, club_info)
    own_id_str = str(own_id)
    termine_je_team: dict[int, list[dict]] = defaultdict(list)
    team_name_by_id: dict[int, str] = {}
    for row in wettkampf_rows:
        heim_id = row["Heimvereinid"].strip()
        gast_id = row["Gastvereinid"].strip()
        if own_id_str not in (heim_id, gast_id):
            continue
        klassenname = row["Klassenname"].strip()
        klassenname_anzeige = rwk.translate_klassenname(translation_map, klassenname)
        disziplin = bestimme_disziplin(klassenname)
        wettkampfnummer = row["Wettkampfnummer"].strip()
        runde = row["Runde"].strip()
        datum = row["Datum"].strip()
        try:
            tag = datetime.strptime(datum, "%d.%m.%Y")
        except ValueError:
            print(f"WARNUNG: Ungueltiges Datum '{datum}' bei Wettkampf {wettkampfnummer}, uebersprungen")
            continue
        (team1_name, team1_id), (team2_name, team2_id) = resolver.resolve_pair(
            klassenname, heim_id, gast_id, wettkampfnummer, datum
        )
        heim_ort = club_info.get(heim_id, {}).get("ort", heim_id)
        gast_ort = club_info.get(gast_id, {}).get("ort", gast_id)
        start = tag.replace(hour=start_uhrzeit[0], minute=start_uhrzeit[1])
        ende = tag.replace(hour=end_uhrzeit[0], minute=end_uhrzeit[1])
        beschreibung = f"{klassenname_anzeige}, Wettkampf {wettkampfnummer}, Runde {runde}"
        if heim_id == own_id_str and team1_id is not None:
            team_name_by_id[team1_id] = team1_name
            termine_je_team[team1_id].append({
                "start": start, "end": ende,
                "saison": saison, "wettkampfnummer": wettkampfnummer,
                "team_id": team1_id, "seite": "heim",
                "titel": f"Wettkampf {disziplin} {gast_ort} Heim",
                "ort": heim_ort,
                "beschreibung": beschreibung,
            })
        if gast_id == own_id_str and team2_id is not None:
            team_name_by_id[team2_id] = team2_name
            termine_je_team[team2_id].append({
                "start": start, "end": ende,
                "saison": saison, "wettkampfnummer": wettkampfnummer,
                "team_id": team2_id, "seite": "gast",
                "titel": f"Wettkampf {disziplin} {heim_ort} Auswärts",
                "ort": heim_ort,
                "beschreibung": beschreibung,
            })
    return termine_je_team, team_name_by_id


def build_ics(termine: list[dict]) -> str:
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        ICS_PRODID,
        "CALSCALE:GREGORIAN",
        *VTIMEZONE_EUROPE_BERLIN.split("\n"),
    ]
    dtstamp = format_dt(datetime.now(timezone.utc)) + "Z"
    for termin in sorted(termine, key=lambda t: t["start"]):
        uid = uuid.uuid5(
            uuid.NAMESPACE_DNS,
            f"disag-rwk:{termin['saison']}:{termin['wettkampfnummer']}:{termin['team_id']}:{termin['seite']}",
        )
        lines += [
            "BEGIN:VEVENT",
            f"UID:{uid}@disag-rwk",
            f"DTSTAMP:{dtstamp}",
            f"DTSTART;TZID=Europe/Berlin:{format_dt(termin['start'])}",
            f"DTEND;TZID=Europe/Berlin:{format_dt(termin['end'])}",
            f"SUMMARY:{ics_escape(termin['titel'])}",
            f"LOCATION:{ics_escape(termin['ort'])}",
            f"DESCRIPTION:{ics_escape(termin['beschreibung'])}",
            "END:VEVENT",
        ]
    lines.append("END:VCALENDAR")
    return "\r\n".join(lines) + "\r\n"


def bereinige_dateien(ausgabe_ordner: Path) -> None:
    for pfad in ausgabe_ordner.glob("*.ics"):
        try:
            ist_export_datei = ICS_PRODID in pfad.read_text(encoding="utf-8")
        except (OSError, UnicodeError):
            continue
        if ist_export_datei:
            pfad.unlink()


def schreibe_dateien(
    termine_je_team: dict[int, list[dict]],
    team_name_by_id: dict[int, str],
    ausgabe_ordner: Path,
) -> int:
    ausgabe_ordner.mkdir(parents=True, exist_ok=True)
    bereinige_dateien(ausgabe_ordner)
    anzahl = 0
    for team_id, termine in termine_je_team.items():
        teamsname = team_name_by_id[team_id]
        pfad = ausgabe_ordner / f"{sichere_dateiname(teamsname)}_{team_id}.ics"
        with open(pfad, "w", encoding="utf-8", newline="") as f:
            f.write(build_ics(termine))
        anzahl += 1
        print(f"ICS-Datei erzeugt: {pfad} ({len(termine)} Termine)")
    return anzahl


def main() -> None:
    args = sys.argv[1:]
    config_path = BASE_DIR / (args[0] if args else "config.ini")
    config = rwk.load_config(config_path)
    rwk.set_debug_level(config.get("Logging", "debug_level", fallback="error"))

    csv_ordner = BASE_DIR / config["Pfade"]["csv_ordner"]
    delimiter = config["Sonstiges"]["csv_trennzeichen"]
    encoding = config["Sonstiges"]["csv_encoding"]

    verein_rows = rwk.read_csv(csv_ordner / config["CSV_Dateien"]["verein_datei"], delimiter, encoding)
    mannschaft_rows = rwk.read_csv(csv_ordner / config["CSV_Dateien"]["mannschaften_datei"], delimiter, encoding)
    wettkampf_rows = rwk.read_csv(csv_ordner / config["CSV_Dateien"]["wettkaempfe_datei"], delimiter, encoding)

    own_id, own_name = rwk.resolve_verein_id(config, verein_rows)
    print(f"Konfigurierter Verein: {own_name} (Vereinsid {own_id})")

    club_info = {
        row["Vereinsid"].strip(): {"name": row["Vereinsname"], "ort": row["Vereinsort"]}
        for row in verein_rows
    }
    translation_map = (
        dict(config.items("Klassen_Uebersetzung")) if config.has_section("Klassen_Uebersetzung") else {}
    )

    saison = config["RWK"]["RWK_Jahr"].strip()
    ausgabe_ordner = (
        BASE_DIR
        / config.get("ICS", "ics_ordner", fallback="RWK_ICS")
        / saison
    )
    start_uhrzeit = parse_uhrzeit(config.get("ICS", "ics_start_uhrzeit", fallback="19:30"))
    end_uhrzeit = parse_uhrzeit(config.get("ICS", "ics_end_uhrzeit", fallback="22:00"))
    if end_uhrzeit <= start_uhrzeit:
        raise ValueError("[ICS] ics_end_uhrzeit muss nach ics_start_uhrzeit liegen")

    teams_by_key = build_teams_by_key(mannschaft_rows, config, club_info, translation_map)
    termine_je_team, team_name_by_id = sammle_termine(
        wettkampf_rows, own_id, saison, teams_by_key, club_info, translation_map, start_uhrzeit, end_uhrzeit
    )

    for (_, vereinsid), teams in teams_by_key.items():
        if vereinsid == str(own_id):
            for _, teamsname, team_id in teams:
                team_name_by_id.setdefault(team_id, teamsname)
                termine_je_team.setdefault(team_id, [])
    anzahl = schreibe_dateien(termine_je_team, team_name_by_id, ausgabe_ordner)
    print(f"{anzahl} ICS-Datei(en) erzeugt in {ausgabe_ordner}.")


if __name__ == "__main__":
    main()
