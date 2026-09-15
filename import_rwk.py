"""Importiert die RWK-CSV-Dateien (Verein, Mannschaften, Schuetzen, Wettkaempfe)
in die Access-Zieldatenbank, wie in Instructions.txt beschrieben.

Aufruf:  .venv\\Scripts\\python.exe import_rwk.py [config.ini]
"""
from __future__ import annotations

import csv
import re
import shutil
import sys
from collections import defaultdict
from configparser import ConfigParser, ExtendedInterpolation
from datetime import datetime
from pathlib import Path

import pyodbc

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE_DIR = Path(__file__).resolve().parent
_truncation_counts: dict[str, int] = defaultdict(int)


def load_config(path: Path) -> ConfigParser:
    config = ConfigParser(interpolation=ExtendedInterpolation())
    with open(path, encoding="utf-8") as f:
        config.read_file(f)
    return config


def read_csv(path: Path, delimiter: str, encoding: str) -> list[dict]:
    with open(path, encoding=encoding, newline="") as f:
        return list(csv.DictReader(f, delimiter=delimiter))


def ensure_target_mdb(config: ConfigParser) -> Path:
    template = BASE_DIR / config["Pfade"]["mdb_vorlage"]
    target = BASE_DIR / config["Pfade"]["mdb_ziel"]
    if not target.exists():
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(template, target)
        template_sdf = template.with_suffix(".sdf")
        if template_sdf.exists():
            shutil.copy2(template_sdf, target.with_suffix(".sdf"))
        print(f"Zieldatenbank angelegt: {target}")
    return target


def connect(mdb_path: Path) -> pyodbc.Connection:
    conn_str = (
        "Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={mdb_path};"
    )
    return pyodbc.connect(conn_str, autocommit=False)


def resolve_verein_id(config: ConfigParser, verein_rows: list[dict]) -> tuple[int, str]:
    pattern = config["RWK"]["Verein"]
    regex = re.compile(pattern, re.IGNORECASE)
    matches = [
        row for row in verein_rows
        if regex.search(row["Vereinsname"]) or regex.search(row["Vereinsort"])
    ]
    if len(matches) != 1:
        found = [(r["Vereinsid"], r["Vereinsname"], r["Vereinsort"]) for r in matches]
        raise ValueError(
            f"Verein-Regex '{pattern}' liefert {len(matches)} Treffer statt genau 1: {found}"
        )
    row = matches[0]
    return int(row["Vereinsid"]), row["Vereinsname"]


def column_lengths(cursor: pyodbc.Cursor, table: str) -> dict[str, int]:
    return {c.column_name.lower(): c.column_size for c in cursor.columns(table=table)}


def find_column(lengths: dict[str, int], *candidates: str) -> str:
    for name in candidates:
        if name.lower() in lengths:
            return name
    raise KeyError(f"Keine der Spalten {candidates} in Tabelle vorhanden")


def truncate(value: str | None, max_len: int | None, context: str) -> str | None:
    if value is None or max_len is None:
        return value
    if len(value) > max_len:
        _truncation_counts[context] += 1
        return value[:max_len]
    return value


def clear_tables(cursor: pyodbc.Cursor) -> None:
    for table in (
        "TeamsShooters",
        "Leaguecompetitions_Competitions",
        "Teams",
        "Shooters",
        "Clubs",
    ):
        cursor.execute(f"DELETE FROM {table}")


def import_clubs(cursor: pyodbc.Cursor, verein_rows: list[dict]) -> None:
    lengths = column_lengths(cursor, "Clubs")
    code_col = find_column(lengths, "clubscode0", "clubscode")
    sql = f"INSERT INTO Clubs (idClubs, clubsname, {code_col}) VALUES (?, ?, ?)"
    for row in verein_rows:
        vereinsid = row["Vereinsid"].strip()
        cursor.execute(
            sql,
            int(vereinsid),
            truncate(row["Vereinsname"], lengths.get("clubsname"), "Clubs.clubsname"),
            truncate(vereinsid, lengths.get(code_col.lower()), f"Clubs.{code_col}"),
        )


def strip_trailing_number(klassenname: str) -> str:
    return re.sub(r"\s*\d+$", "", klassenname).strip()


def import_teams(
    cursor: pyodbc.Cursor, mannschaft_rows: list[dict], config: ConfigParser
) -> dict[tuple[str, str], list[tuple[int, str]]]:
    lengths = column_lengths(cursor, "Teams")
    name_format = config["Teams"]["teamsname_format"]
    nameshort_format = config["Teams"]["teamsnameshort_format"]
    sql = (
        "INSERT INTO Teams (idTeams, fidClubs, teamsname, teamsnameshort) "
        "VALUES (?, ?, ?, ?)"
    )
    teams_by_key: dict[tuple[str, str], list[tuple[int, str]]] = defaultdict(list)
    for row in mannschaft_rows:
        vereinsid = row["Vereinsid"].strip()
        klassenname = row["Klassenname"].strip()
        nummer = int(row["Mannschaftsnummer"])
        format_vars = {
            "klassenname": klassenname,
            "klasse_ohne_zahl": strip_trailing_number(klassenname),
            "nummer": nummer,
            "vereinsid": vereinsid,
            "mannschaftsid": row["Mannschaftsid"].strip(),
        }
        teamsname = truncate(
            name_format.format(**format_vars), lengths.get("teamsname"), "Teams.teamsname"
        )
        teamsnameshort = truncate(
            nameshort_format.format(**format_vars),
            lengths.get("teamsnameshort"),
            "Teams.teamsnameshort",
        )
        cursor.execute(
            sql, int(row["Mannschaftsid"]), int(vereinsid), teamsname, teamsnameshort
        )
        teams_by_key[(klassenname, vereinsid)].append((nummer, teamsname))
    for key, candidates in teams_by_key.items():
        candidates.sort(key=lambda c: c[0])
    return teams_by_key


def import_shooters(
    cursor: pyodbc.Cursor, schuetzen_rows: list[dict], id_offset: int
) -> tuple[dict[tuple[str, str], int], dict[str, int]]:
    lengths = column_lengths(cursor, "Shooters")
    sql = (
        "INSERT INTO Shooters "
        "(idShooters, firstname, lastname, fidClubs, birthyear, sex, internid, identification) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
    )
    lookup_by_club: dict[tuple[str, str], int] = {}
    lookup_any: dict[str, int] = {}
    for row in schuetzen_rows:
        schuetzenid = row["Schuetzenid"].strip()
        vereinsids = [v.strip() for v in row["Vereinsids"].split(",") if v.strip()]
        for i, vereinsid in enumerate(vereinsids):
            id_final = int(schuetzenid) + id_offset * i
            cursor.execute(
                sql,
                id_final,
                truncate(row["Vorname"], lengths.get("firstname"), "Shooters.firstname"),
                truncate(row["Nachname"], lengths.get("lastname"), "Shooters.lastname"),
                int(vereinsid),
                1900,
                0,
                schuetzenid,
                truncate(schuetzenid, lengths.get("identification"), "Shooters.identification"),
            )
            lookup_by_club[(schuetzenid, vereinsid)] = id_final
            lookup_any.setdefault(schuetzenid, id_final)
    return lookup_by_club, lookup_any


def import_teams_shooters(
    cursor: pyodbc.Cursor,
    mannschaft_rows: list[dict],
    lookup_by_club: dict[tuple[str, str], int],
    lookup_any: dict[str, int],
) -> None:
    sql = "INSERT INTO TeamsShooters (fidTeams, fidShooters) VALUES (?, ?)"
    for row in mannschaft_rows:
        vereinsid = row["Vereinsid"].strip()
        idteams = int(row["Mannschaftsid"])
        for col in ("Schuetze1", "Schuetze2", "Schuetze3", "Schuetze4", "Schuetze5", "Schuetze6"):
            schuetzenid = (row.get(col) or "").strip()
            if not schuetzenid:
                continue
            idshooters = lookup_by_club.get((schuetzenid, vereinsid))
            if idshooters is None:
                idshooters = lookup_any.get(schuetzenid)
                if idshooters is not None:
                    print(
                        f"WARNUNG: Schuetze {schuetzenid} nicht bei Verein {vereinsid} "
                        f"gefunden, verwende Eintrag aus anderem Verein"
                    )
            if idshooters is None:
                print(f"WARNUNG: Schuetze {schuetzenid} (Mannschaft {idteams}) nicht gefunden")
                continue
            cursor.execute(sql, idteams, idshooters)


class TeamResolver:
    """Loest team1/team2-Namen auf; behandelt Vereine mit mehreren Mannschaften
    in derselben Klasse (u.a. Vereins-Derbys) per Rotation ueber Hin-/Rueckrunde."""

    def __init__(self, teams_by_key: dict[tuple[str, str], list[tuple[int, str]]]):
        self.teams_by_key = teams_by_key
        self.counters: dict[tuple[str, str], int] = defaultdict(int)

    def resolve_pair(
        self, klassenname: str, heim_id: str, gast_id: str, wettkampfnummer: str
    ) -> tuple[str | None, str | None]:
        heim_cands = self.teams_by_key.get((klassenname, heim_id), [])
        gast_cands = self.teams_by_key.get((klassenname, gast_id), [])
        if not heim_cands:
            print(f"WARNUNG: Keine Heimmannschaft fuer Klasse '{klassenname}' Verein {heim_id} (Wettkampf {wettkampfnummer})")
        if not gast_cands:
            print(f"WARNUNG: Keine Gastmannschaft fuer Klasse '{klassenname}' Verein {gast_id} (Wettkampf {wettkampfnummer})")

        if heim_id == gast_id and len(heim_cands) > 1:
            key = (klassenname, heim_id)
            first_home = self.counters[key] % 2 == 0
            self.counters[key] += 1
            a, b = heim_cands[0][1], heim_cands[1][1]
            return (a, b) if first_home else (b, a)

        return (
            self._pick(heim_cands, klassenname, heim_id),
            self._pick(gast_cands, klassenname, gast_id),
        )

    def _pick(self, candidates: list[tuple[int, str]], klassenname: str, vereinsid: str) -> str | None:
        if not candidates:
            return None
        if len(candidates) == 1:
            return candidates[0][1]
        key = (klassenname, vereinsid)
        idx = self.counters[key] % len(candidates)
        self.counters[key] += 1
        print(
            f"WARNUNG: Mehrdeutige Mannschaft fuer Klasse '{klassenname}' / Verein {vereinsid} "
            f"- verwende Mannschaft #{candidates[idx][0]} (Rotation {idx + 1}/{len(candidates)})"
        )
        return candidates[idx][1]


def import_leaguecompetitions(
    cursor: pyodbc.Cursor,
    wettkampf_rows: list[dict],
    vereinsid: int,
    teams_by_key: dict[tuple[str, str], list[tuple[int, str]]],
) -> None:
    lengths = column_lengths(cursor, "Leaguecompetitions_Competitions")
    resolver = TeamResolver(teams_by_key)
    sql = (
        "INSERT INTO Leaguecompetitions_Competitions "
        "(idLeaguecompetitions_Competitions, [name], team1, team2, [date], [type], additional_info) "
        "VALUES (?, ?, ?, ?, ?, ?, ?)"
    )
    vereinsid_str = str(vereinsid)
    for row in wettkampf_rows:
        if row["Heimvereinid"].strip() != vereinsid_str:
            continue
        klassenname = row["Klassenname"].strip()
        wettkampfnummer = row["Wettkampfnummer"].strip()
        runde = row["Runde"].strip()
        datum = row["Datum"].strip()
        name = truncate(
            f"{klassenname}, {wettkampfnummer}-{runde}", lengths.get("name"),
            "Leaguecompetitions_Competitions.name",
        )
        team1, team2 = resolver.resolve_pair(
            klassenname, row["Heimvereinid"].strip(), row["Gastvereinid"].strip(), wettkampfnummer
        )
        team1 = truncate(team1, lengths.get("team1"), "Leaguecompetitions_Competitions.team1")
        team2 = truncate(team2, lengths.get("team2"), "Leaguecompetitions_Competitions.team2")
        date = datetime.strptime(datum, "%d.%m.%Y").replace(hour=20, minute=0, second=0)
        additional_info = truncate(
            datum, lengths.get("additional_info"), "Leaguecompetitions_Competitions.additional_info"
        )
        cursor.execute(sql, int(wettkampfnummer), name, team1, team2, date, 0, additional_info)


def main() -> None:
    config_path = BASE_DIR / (sys.argv[1] if len(sys.argv) > 1 else "config.ini")
    config = load_config(config_path)

    csv_ordner = BASE_DIR / config["Pfade"]["csv_ordner"]
    delimiter = config["Sonstiges"]["csv_trennzeichen"]
    encoding = config["Sonstiges"]["csv_encoding"]
    id_offset = int(config["Sonstiges"]["schuetzen_id_offset"])

    verein_rows = read_csv(csv_ordner / config["CSV_Dateien"]["verein_datei"], delimiter, encoding)
    mannschaft_rows = read_csv(csv_ordner / config["CSV_Dateien"]["mannschaften_datei"], delimiter, encoding)
    schuetzen_rows = read_csv(csv_ordner / config["CSV_Dateien"]["schuetzen_datei"], delimiter, encoding)
    wettkampf_rows = read_csv(csv_ordner / config["CSV_Dateien"]["wettkaempfe_datei"], delimiter, encoding)

    vereinsid, vereinsname = resolve_verein_id(config, verein_rows)
    print(f"Konfigurierter Verein: {vereinsname} (Vereinsid {vereinsid})")

    mdb_path = ensure_target_mdb(config)
    conn = connect(mdb_path)
    try:
        cursor = conn.cursor()
        clear_tables(cursor)
        import_clubs(cursor, verein_rows)
        teams_by_key = import_teams(cursor, mannschaft_rows, config)
        lookup_by_club, lookup_any = import_shooters(cursor, schuetzen_rows, id_offset)
        import_teams_shooters(cursor, mannschaft_rows, lookup_by_club, lookup_any)
        import_leaguecompetitions(cursor, wettkampf_rows, vereinsid, teams_by_key)
        conn.commit()
        for context, count in _truncation_counts.items():
            print(f"WARNUNG: {count}x gekuerzt, da zu lang fuer Feld: {context}")
        print("Import erfolgreich abgeschlossen.")
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


if __name__ == "__main__":
    main()
