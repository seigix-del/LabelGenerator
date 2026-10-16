"""Regenerates access/sql/*.sql from the SQL embedded in modBuildDatabase.bas.

Run after editing the builder:  python access/tools/extract_sql.py
"""
import pathlib
import re

ROOT = pathlib.Path(__file__).resolve().parent.parent
BAS = ROOT / "vba" / "modBuildDatabase.bas"
LIT = re.compile(r'"((?:[^"]|"")*)"')


def statements():
    lines = BAS.read_text(encoding="cp1252").splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m = re.match(r"(RunDDL|SaveQuery)\b", line)
        if not m:
            i += 1
            continue
        chunk = [line]
        while chunk[-1].endswith("_"):
            i += 1
            chunk.append(lines[i].strip())
        lits = [x.replace('""', '"') for x in LIT.findall(" ".join(chunk))]
        yield m.group(1), lits[0], "".join(lits[1:])
        i += 1


def main():
    ddl, queries = [], []
    for kind, name, sql in statements():
        if kind == "RunDDL":
            ddl.append(sql.replace(", CONSTRAINT", ",\n    CONSTRAINT") + ";")
        else:
            queries.append(f"-- {name}\n{sql};")
    header = "-- Generated from access/vba/modBuildDatabase.bas by access/tools/extract_sql.py - do not edit.\n"
    (ROOT / "sql" / "01_create_tables.sql").write_text(
        header + "-- Access DDL, ANSI-92 mode (run through ADO: CurrentProject.Connection.Execute).\n\n"
        + "\n\n".join(ddl) + "\n", encoding="utf-8")
    (ROOT / "sql" / "02_queries.sql").write_text(
        header + "-- Saved Access queries (DAO / ANSI-89). Name of each query is on the line above it.\n\n"
        + "\n\n".join(queries) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
