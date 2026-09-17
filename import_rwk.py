"""Importiert die RWK-CSV-Dateien (Verein, Mannschaften, Schuetzen, Wettkaempfe)
in die Access-Zieldatenbank, wie in Instructions.txt beschrieben.

Aufruf:  .venv\\Scripts\\python.exe import_rwk.py [config.ini]
"""
from __future__ import annotations

import csv
import re
import shutil
import sys
import uuid
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
    config.optionxform = str  # Gross-/Kleinschreibung der Keys erhalten (z.B. Klassennamen)
    with open(path, encoding="utf-8") as f:
        config.read_file(f)
    return config


def split_klassenname(klassenname: str) -> tuple[str, str]:
    """Trennt den Klassennamen in Stamm (ohne Zahl) und die abschliessende Zahl (falls vorhanden)."""
    match = re.search(r"\s*(\d+)$", klassenname)
    if not match:
        return klassenname.strip(), ""
    return klassenname[: match.start()].strip(), match.group(1)


def translate_klassenname(translation_map: dict[str, str], klassenname: str) -> str:
    """Uebersetzt nur den Namensstamm (ohne die abschliessende Mannschaftsnummer)."""
    stamm, nummer = split_klassenname(klassenname)
    uebersetzt = translation_map.get(stamm, stamm)
    return f"{uebersetzt} {nummer}".strip() if nummer else uebersetzt


def ensure_klassen_uebersetzung(
    config: ConfigParser, config_path: Path, klassennamen: list[str]
) -> dict[str, str]:
    """Ergaenzt config.ini um eine (optionale) Uebersetzungstabelle fuer alle in
    Wettkaempfe_<Jahr>.csv gefundenen Klassennamen, sofern noch nicht vorhanden."""
    section = "Klassen_Uebersetzung"
    existing = dict(config.items(section)) if config.has_section(section) else {}
    missing = [name for name in klassennamen if name not in existing]
    if not missing:
        return existing
    if not config.has_section(section):
        config.add_section(section)
    lines = []
    if not existing:
        lines += [
            "",
            f"[{section}]",
            "; Optionale Uebersetzung/Klartext fuer Klassennamen aus Wettkaempfe_<Jahr>.csv.",
            "; Format: <Klassenname aus CSV> = <Anzeigename>",
        ]
    for name in missing:
        config.set(section, name, name)
        lines.append(f"{name} = {name}")
    with open(config_path, "a", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print(f"Abschnitt [{section}]: {len(missing)} neue Klassennamen ergaenzt in {config_path.name}.")
    return dict(config.items(section))


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
        "Leaguecompetitions_ShootersShots",
        "Leaguecompetitions_Shooters",
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
    cursor: pyodbc.Cursor,
    mannschaft_rows: list[dict],
    config: ConfigParser,
    club_info: dict[str, dict[str, str]],
    translation_map: dict[str, str],
) -> dict[tuple[str, str], list[tuple[int, str, int]]]:
    lengths = column_lengths(cursor, "Teams")
    name_format = config["Teams"]["teamsname_format"]
    nameshort_format = config["Teams"]["teamsnameshort_format"]
    sql = (
        "INSERT INTO Teams (idTeams, fidClubs, teamsname, teamsnameshort) "
        "VALUES (?, ?, ?, ?)"
    )
    teams_by_key: dict[tuple[str, str], list[tuple[int, str, int]]] = defaultdict(list)
    for row in mannschaft_rows:
        vereinsid = row["Vereinsid"].strip()
        klassenname = row["Klassenname"].strip()
        klassenname_anzeige = translate_klassenname(translation_map, klassenname)
        nummer = int(row["Mannschaftsnummer"])
        idteams = int(row["Mannschaftsid"])
        club = club_info.get(vereinsid, {})
        format_vars = {
            "klassenname": klassenname_anzeige,
            "klasse_ohne_zahl": strip_trailing_number(klassenname_anzeige),
            "nummer": nummer,
            "vereinsid": vereinsid,
            "mannschaftsid": row["Mannschaftsid"].strip(),
            "vereinsname": club.get("name", ""),
            "vereinsort": club.get("ort", ""),
        }
        teamsname = truncate(
            name_format.format(**format_vars), lengths.get("teamsname"), "Teams.teamsname"
        )
        teamsnameshort = truncate(
            nameshort_format.format(**format_vars),
            lengths.get("teamsnameshort"),
            "Teams.teamsnameshort",
        )
        cursor.execute(sql, idteams, int(vereinsid), teamsname, teamsnameshort)
        teams_by_key[(klassenname, vereinsid)].append((nummer, teamsname, idteams))
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


def resolve_shooter_id(
    schuetzenid: str,
    vereinsid: str,
    lookup_by_club: dict[tuple[str, str], int],
    lookup_any: dict[str, int],
) -> int | None:
    idshooters = lookup_by_club.get((schuetzenid, vereinsid))
    if idshooters is None:
        idshooters = lookup_any.get(schuetzenid)
        if idshooters is not None:
            print(
                f"WARNUNG: Schuetze {schuetzenid} nicht bei Verein {vereinsid} "
                f"gefunden, verwende Eintrag aus anderem Verein"
            )
    return idshooters


def roster_schuetzen(row: dict) -> list[str]:
    columns = ("Schuetze1", "Schuetze2", "Schuetze3", "Schuetze4", "Schuetze5", "Schuetze6")
    return [s for s in ((row.get(c) or "").strip() for c in columns) if s]


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
        for schuetzenid in roster_schuetzen(row):
            idshooters = resolve_shooter_id(schuetzenid, vereinsid, lookup_by_club, lookup_any)
            if idshooters is None:
                print(f"WARNUNG: Schuetze {schuetzenid} (Mannschaft {idteams}) nicht gefunden")
                continue
            cursor.execute(sql, idteams, idshooters)


class TeamResolver:
    """Loest team1/team2 (Name + idTeams) auf; hat ein Verein mehrere Mannschaften
    in derselben Klasse (u.a. bei Vereins-Derbys), wird der Benutzer beim ersten
    betroffenen Wettkampf interaktiv gefragt, welche Mannschaft (1, 2, ...) gemeint
    ist. Bei weiteren Wettkaempfen mit derselben Klasse/demselben Verein wird
    automatisch abwechselnd die jeweils andere Mannschaft verwendet."""

    def __init__(self, teams_by_key: dict[tuple[str, str], list[tuple[int, str, int]]]):
        self.teams_by_key = teams_by_key
        self.last_choice: dict[tuple[str, str], tuple[int, str, int]] = {}

    def resolve_pair(
        self, klassenname: str, heim_id: str, gast_id: str, wettkampfnummer: str, datum: str
    ) -> tuple[tuple[str | None, int | None], tuple[str | None, int | None]]:
        heim_cands = self.teams_by_key.get((klassenname, heim_id), [])
        gast_cands = self.teams_by_key.get((klassenname, gast_id), [])
        if not heim_cands:
            print(f"WARNUNG: Keine Heimmannschaft fuer Klasse '{klassenname}' Verein {heim_id} (Wettkampf {wettkampfnummer})")
        if not gast_cands:
            print(f"WARNUNG: Keine Gastmannschaft fuer Klasse '{klassenname}' Verein {gast_id} (Wettkampf {wettkampfnummer})")

        if heim_id == gast_id and len(heim_cands) > 1:
            key = (klassenname, heim_id)
            a, b = heim_cands[0], heim_cands[1]
            previous = self.last_choice.get(key)
            if previous is not None:
                chosen_first = self._other((a, b), previous)
            else:
                chosen_first = self._ask(
                    f"Wettkampf {wettkampfnummer} am {datum}, Klasse '{klassenname}': Vereins-Derby bei "
                    f"Verein {heim_id} - welche Mannschaft ist die erste (Heim)?",
                    (a, b),
                )
            self.last_choice[key] = chosen_first
            a, b = (a, b) if chosen_first is a else (b, a)
            return (a[1], a[2]), (b[1], b[2])

        return (
            self._pick(heim_cands, klassenname, heim_id, gast_id, "Heim", wettkampfnummer, datum),
            self._pick(gast_cands, klassenname, gast_id, heim_id, "Gast", wettkampfnummer, datum),
        )

    def _pick(
        self,
        candidates: list[tuple[int, str, int]],
        klassenname: str,
        vereinsid: str,
        gegner_id: str,
        rolle: str,
        wettkampfnummer: str,
        datum: str,
    ) -> tuple[str | None, int | None]:
        if not candidates:
            return None, None
        if len(candidates) == 1:
            return candidates[0][1], candidates[0][2]
        key = (klassenname, vereinsid)
        previous = self.last_choice.get(key)
        if previous is not None:
            chosen = self._other(candidates, previous)
        else:
            chosen = self._ask(
                f"Wettkampf {wettkampfnummer} am {datum}, Klasse '{klassenname}': Verein {vereinsid} hat "
                f"mehrere Mannschaften und spielt als {rolle} gegen Verein {gegner_id}. "
                f"Welche Mannschaft ist die erste?",
                candidates,
            )
        self.last_choice[key] = chosen
        return chosen[1], chosen[2]

    @staticmethod
    def _other(
        candidates: tuple[tuple[int, str, int], ...] | list[tuple[int, str, int]],
        previous: tuple[int, str, int],
    ) -> tuple[int, str, int]:
        idx = candidates.index(previous)
        return candidates[(idx + 1) % len(candidates)]

    @staticmethod
    def _ask(
        frage: str, candidates: tuple[tuple[int, str, int], ...] | list[tuple[int, str, int]]
    ) -> tuple[int, str, int]:
        options = ", ".join(str(c[0]) for c in candidates)
        prompt = f"{frage} [{options}]: "
        while True:
            answer = input(prompt).strip()
            for cand in candidates:
                if str(cand[0]) == answer:
                    return cand
            print(f"Ungueltige Eingabe. Bitte eine der Nummern eingeben: {options}")


def insert_competition_shooters(
    cursor: pyodbc.Cursor,
    id_lc: int,
    team_idteams: int | None,
    team_slot: int,
    mannschaft_by_id: dict[int, dict],
    lookup_by_club: dict[tuple[str, str], int],
    lookup_any: dict[str, int],
) -> None:
    if team_idteams is None:
        return
    row = mannschaft_by_id.get(team_idteams)
    if row is None:
        return
    vereinsid = row["Vereinsid"].strip()
    sql = (
        "INSERT INTO Leaguecompetitions_Shooters "
        "(fidShooters, assessed, team, fidLeaguecompetitions_Competitions, fidRanges, position, fidDisciplines, resultid) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
    )
    for schuetzenid in roster_schuetzen(row):
        idshooters = resolve_shooter_id(schuetzenid, vereinsid, lookup_by_club, lookup_any)
        if idshooters is None:
            print(f"WARNUNG: Schuetze {schuetzenid} (Wettkampf {id_lc}) nicht gefunden")
            continue
        cursor.execute(
            sql, idshooters, 1, team_slot, id_lc, 0, 0, 0, str(uuid.uuid4()).upper()
        )


def import_leaguecompetitions(
    cursor: pyodbc.Cursor,
    wettkampf_rows: list[dict],
    vereinsid: int,
    teams_by_key: dict[tuple[str, str], list[tuple[int, str, int]]],
    mannschaft_by_id: dict[int, dict],
    lookup_by_club: dict[tuple[str, str], int],
    lookup_any: dict[str, int],
    translation_map: dict[str, str],
    club_info: dict[str, dict[str, str]],
) -> list[dict]:
    lengths = column_lengths(cursor, "Leaguecompetitions_Competitions")
    resolver = TeamResolver(teams_by_key)
    sql = (
        "INSERT INTO Leaguecompetitions_Competitions "
        "(idLeaguecompetitions_Competitions, [name], team1, team2, [date], [type], additional_info) "
        "VALUES (?, ?, ?, ?, ?, ?, ?)"
    )
    vereinsid_str = str(vereinsid)
    heimwettkaempfe: list[dict] = []
    for row in wettkampf_rows:
        if row["Heimvereinid"].strip() != vereinsid_str:
            continue
        klassenname = row["Klassenname"].strip()
        klassenname_anzeige = translate_klassenname(translation_map, klassenname)
        wettkampfnummer = row["Wettkampfnummer"].strip()
        runde = row["Runde"].strip()
        datum = row["Datum"].strip()
        name = truncate(
            f"{klassenname_anzeige}, {wettkampfnummer}-{runde}", lengths.get("name"),
            "Leaguecompetitions_Competitions.name",
        )
        (team1_name, team1_id), (team2_name, team2_id) = resolver.resolve_pair(
            klassenname, row["Heimvereinid"].strip(), row["Gastvereinid"].strip(), wettkampfnummer, datum
        )
        team1 = truncate(team1_name, lengths.get("team1"), "Leaguecompetitions_Competitions.team1")
        team2 = truncate(team2_name, lengths.get("team2"), "Leaguecompetitions_Competitions.team2")
        date = datetime.strptime(datum, "%d.%m.%Y").replace(hour=20, minute=0, second=0)
        additional_info = truncate(
            datum, lengths.get("additional_info"), "Leaguecompetitions_Competitions.additional_info"
        )
        id_lc = int(wettkampfnummer)
        cursor.execute(sql, id_lc, name, team1, team2, date, 0, additional_info)
        insert_competition_shooters(cursor, id_lc, team1_id, 1, mannschaft_by_id, lookup_by_club, lookup_any)
        insert_competition_shooters(cursor, id_lc, team2_id, 2, mannschaft_by_id, lookup_by_club, lookup_any)
        heim_ort = club_info.get(row["Heimvereinid"].strip(), {}).get("ort", row["Heimvereinid"].strip())
        gast_ort = club_info.get(row["Gastvereinid"].strip(), {}).get("ort", row["Gastvereinid"].strip())
        heimwettkaempfe.append({
            "Wettkampfnummer": wettkampfnummer,
            "Runde": runde,
            "Datum": datum,
            "Klassenname": klassenname_anzeige,
            "Heimort": heim_ort,
            "Gastort": gast_ort,
            "Team1": team1_name or "",
            "Team2": team2_name or "",
        })
    return heimwettkaempfe


def export_heimwettkaempfe(config: ConfigParser, rows: list[dict]) -> None:
    if "Export" not in config:
        return
    export_cfg = config["Export"]
    ordner = export_cfg.get("heimwettkaempfe_ordner")
    dateiname = export_cfg.get("heimwettkaempfe_datei")
    if not ordner or not dateiname:
        return
    delimiter = config["Sonstiges"]["csv_trennzeichen"]
    encoding = config["Sonstiges"]["csv_encoding"]
    out_dir = BASE_DIR / ordner
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / dateiname
    fieldnames = [
        "Wettkampfnummer", "Runde", "Datum", "Klassenname",
        "Heimort", "Gastort", "Team1", "Team2",
    ]
    with open(out_path, "w", encoding=encoding, newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=delimiter)
        writer.writeheader()
        writer.writerows(rows)
    print(f"Heimwettkaempfe exportiert: {out_path} ({len(rows)} Zeilen)")


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

    klassen_stamme = sorted({
        split_klassenname(row["Klassenname"].strip())[0]
        for row in wettkampf_rows if row["Klassenname"].strip()
    })
    translation_map = ensure_klassen_uebersetzung(config, config_path, klassen_stamme)

    mdb_path = ensure_target_mdb(config)
    conn = connect(mdb_path)
    try:
        cursor = conn.cursor()
        clear_tables(cursor)
        import_clubs(cursor, verein_rows)
        club_info = {
            row["Vereinsid"].strip(): {"name": row["Vereinsname"], "ort": row["Vereinsort"]}
            for row in verein_rows
        }
        teams_by_key = import_teams(cursor, mannschaft_rows, config, club_info, translation_map)
        lookup_by_club, lookup_any = import_shooters(cursor, schuetzen_rows, id_offset)
        import_teams_shooters(cursor, mannschaft_rows, lookup_by_club, lookup_any)
        mannschaft_by_id = {int(row["Mannschaftsid"]): row for row in mannschaft_rows}
        heimwettkaempfe = import_leaguecompetitions(
            cursor, wettkampf_rows, vereinsid, teams_by_key,
            mannschaft_by_id, lookup_by_club, lookup_any, translation_map, club_info,
        )
        conn.commit()
        export_heimwettkaempfe(config, heimwettkaempfe)
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
