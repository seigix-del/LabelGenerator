"""Builds the Power BI project (PBIP, TMDL semantic model) from powerbi/queries/*.pq.

    python powerbi/tools/build_pbip.py

Writes powerbi/InventoryTracker.pbip, InventoryTracker.SemanticModel/,
InventoryTracker.Report/ and powerbi/dax/measures.dax. The .pq files are the
single source of truth; re-run this after editing them.
"""
import json
import pathlib
import shutil
import uuid

ROOT = pathlib.Path(__file__).resolve().parent.parent
QUERIES = ROOT / "queries"
NAME = "InventoryTracker"
NS = uuid.UUID("6f1c2a52-5d0e-4c55-9a43-1d3f0a7b9e10")

# Not loaded: parameters, config, functions and staging queries -> TMDL expressions.
EXPRESSIONS = [
    ("MasterFolder", "Text"), ("StatusFolder", "Text"),
    ("cfgInEvents", "List"), ("cfgGrades", "List"), ("cfgHeaderAliases", "Table"),
    ("fnCleanText", "Function"), ("fnNormKey", "Function"), ("fnNormWaybill", "Function"),
    ("fnHeaderLogical", "Function"), ("fnParseGradeLoss", "Function"), ("fnCustomerCode", "Function"),
    ("fnReadWorkbook", "Function"), ("fnExcelFiles", "Function"),
    ("MasterClean", "Table"), ("StatusAll", "Table"), ("StatusLatest", "Table"),
]

# Loaded tables: name -> columns (name, dataType, summarizeBy, extra lines)
TABLES = {
    "Shipments": [
        ("Waybill", "string", "none", []), ("Reference", "string", "none", []),
        ("SecondaryRef", "string", "none", []), ("Contents", "string", "none", []),
        ("Grade", "string", "none", []), ("RowType", "string", "none", []),
        ("GQty", "int64", "sum", ["isHidden"]), ("LossQty", "int64", "sum", ["isHidden"]),
        ("LastEventLabel", "string", "none", []), ("HasStatus", "boolean", "none", []),
        ("InWarehouse", "boolean", "none", []), ("Location", "string", "none", []),
        ("City", "string", "none", []), ("Consignee", "string", "none", []),
        ("CustomerCode", "string", "none", []), ("MasterFile", "string", "none", []),
        ("MasterFileDate", "dateTime", "none", ["formatString: General Date"]),
        ("StatusFile", "string", "none", []),
        ("StatusFileDate", "dateTime", "none", ["formatString: General Date"]),
    ],
    "Unmatched": [
        ("Waybill", "string", "none", []), ("LastEvent", "string", "none", []),
        ("CityName", "string", "none", []), ("ConsigneeName", "string", "none", []),
        ("SourceFile", "string", "none", []),
        ("FileDate", "dateTime", "none", ["formatString: General Date"]),
    ],
    "StatusHistory": [
        ("Waybill", "string", "none", []), ("LastEvent", "string", "none", []),
        ("SourceFile", "string", "none", []),
        ("FileDate", "dateTime", "none", ["formatString: General Date"]),
        ("Location", "string", "none", []),
    ],
    "Grades": [
        ("Grade", "string", "none", ["sortByColumn: SortOrder"]),
        ("SortOrder", "int64", "none", ["isHidden"]),
    ],
}

# Measures (all on Shipments): name, DAX, format, display folder
MEASURES = [
    ("Total Waybills", "COUNTROWS(Shipments)", "#,0", "Inventory"),
    ("IN Waybills", "CALCULATE(COUNTROWS(Shipments), Shipments[InWarehouse] = TRUE())", "#,0", "Inventory"),
    ("OUT Waybills", "CALCULATE(COUNTROWS(Shipments), Shipments[InWarehouse] = FALSE())", "#,0", "Inventory"),
    ("% IN", "DIVIDE([IN Waybills], [Total Waybills])", "0.0%", "Inventory"),
    ("Unmatched Waybills", "COUNTROWS(Unmatched) + 0", "#,0", "Inventory"),
    ("No Grade Waybills", "CALCULATE(COUNTROWS(Shipments), ISBLANK(Shipments[Grade])) + 0", "#,0", "Inventory"),
    ("Last Status File Date", "MAX(Shipments[StatusFileDate])", "yyyy-mm-dd hh:nn", "Inventory"),
    ("G Total", "SUM(Shipments[GQty])", "#,0", "Grade stock"),
    ("G IN", "CALCULATE(SUM(Shipments[GQty]), Shipments[InWarehouse] = TRUE())", "#,0", "Grade stock"),
    ("G OUT", "CALCULATE(SUM(Shipments[GQty]), Shipments[InWarehouse] = FALSE())", "#,0", "Grade stock"),
    ("G Remaining", "[G Total] - [G OUT]", "#,0", "Grade stock"),
    ("L Total", "SUM(Shipments[LossQty])", "#,0", "Grade stock"),
    ("L IN", "CALCULATE(SUM(Shipments[LossQty]), Shipments[InWarehouse] = TRUE())", "#,0", "Grade stock"),
    ("L OUT", "CALCULATE(SUM(Shipments[LossQty]), Shipments[InWarehouse] = FALSE())", "#,0", "Grade stock"),
    ("L Remaining", "[L Total] - [L OUT]", "#,0", "Grade stock"),
]

RELATIONSHIPS = [("Shipments", "Grade", "Grades", "Grade")]


def tag(*parts):
    return str(uuid.uuid5(NS, "/".join(parts)))


def q(name):
    return f"'{name}'" if any(ch in name for ch in " %-") else name


def read_query(name):
    """Query text without its leading // comment lines (kept in the .pq file)."""
    lines = (QUERIES / f"{name}.pq").read_text(encoding="utf-8").strip("\n").splitlines()
    while lines and lines[0].lstrip().startswith("//"):
        lines.pop(0)
    return "\n".join(lines)


def indent(text, tabs):
    pad = "\t" * tabs
    return "\n".join(pad + line if line.strip() else "" for line in text.splitlines())


def expressions_tmdl():
    out = []
    for name, kind in EXPRESSIONS:
        body = read_query(name)
        if kind == "Text":           # parameter: single line with meta record
            out.append(f"expression {name} = {body}\n\tlineageTag: {tag('expr', name)}\n\n"
                       f"\tannotation PBI_ResultType = Text\n")
        else:
            out.append(f"expression {name} =\n{indent(body, 2)}\n\tlineageTag: {tag('expr', name)}\n\n"
                       f"\tannotation PBI_ResultType = {kind}\n")
    return "\n".join(out)


def table_tmdl(name, columns):
    lines = [f"table {name}", f"\tlineageTag: {tag('table', name)}", ""]
    if name == "Shipments":
        for m, dax, fmt, folder in MEASURES:
            lines += [f"\tmeasure {q(m)} = {dax}", f"\t\tformatString: {fmt}",
                      f"\t\tdisplayFolder: {folder}", f"\t\tlineageTag: {tag('measure', m)}", ""]
    for col, dtype, summ, extra in columns:
        lines += [f"\tcolumn {q(col)}", f"\t\tdataType: {dtype}"]
        lines += [f"\t\t{e}" for e in extra]
        lines += [f"\t\tlineageTag: {tag('column', name, col)}", f"\t\tsummarizeBy: {summ}",
                  f"\t\tsourceColumn: {col}", ""]
    lines += [f"\tpartition {name} = m", "\t\tmode: import", "\t\tsource =",
              indent(read_query(name), 4), ""]
    return "\n".join(lines) + "\n"


def write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def main():
    sm = ROOT / f"{NAME}.SemanticModel"
    rp = ROOT / f"{NAME}.Report"
    for d in (sm, rp):
        if d.exists():
            shutil.rmtree(d)

    write(ROOT / f"{NAME}.pbip", json.dumps({
        "version": "1.0",
        "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
        "settings": {"enableAutoRecovery": True},
    }, indent=2) + "\n")

    # --- semantic model (TMDL) ---
    write(sm / "definition.pbism", json.dumps({"version": "4.0", "settings": {}}, indent=2) + "\n")
    write(sm / "definition" / "database.tmdl", "database\n\tcompatibilityLevel: 1567\n")
    model = ["model Model", "\tculture: en-US", "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
             "\tdiscourageImplicitMeasures", "\tsourceQueryCulture: en-US", "\tdataAccessOptions",
             "\t\tlegacyRedirects", "\t\treturnErrorValuesAsNull", "",
             "annotation PBI_QueryOrder = " + json.dumps([n for n, _ in EXPRESSIONS] + list(TABLES)), ""]
    model += [f"ref table {t}" for t in TABLES]
    write(sm / "definition" / "model.tmdl", "\n".join(model) + "\n")
    write(sm / "definition" / "expressions.tmdl", expressions_tmdl())
    for name, cols in TABLES.items():
        write(sm / "definition" / "tables" / f"{name}.tmdl", table_tmdl(name, cols))
    rel = []
    for ft, fc, tt, tc in RELATIONSHIPS:
        rel += [f"relationship {tag('rel', ft, fc, tt, tc)}", f"\tfromColumn: {ft}.{fc}",
                f"\ttoColumn: {tt}.{tc}", ""]
    write(sm / "definition" / "relationships.tmdl", "\n".join(rel))

    # --- report: one blank page per planned page; visuals are added in Desktop ---
    pages = ["Dashboard", "Grade stock", "Consignee x grade", "Shipments", "Unmatched"]
    sections = [{
        "config": "{}", "displayName": p, "displayOption": 1, "filters": "[]",
        "height": 720.0, "width": 1280.0, "name": f"ReportSection{i}", "ordinal": i,
        "visualContainers": [],
    } for i, p in enumerate(pages)]
    write(rp / "definition.pbir", json.dumps({
        "version": "1.0",
        "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}, "byConnection": None},
    }, indent=2) + "\n")
    write(rp / "report.json", json.dumps({
        "config": json.dumps({"version": "5.43", "themeCollection": {}, "activeSectionIndex": 0}),
        "layoutOptimization": 0,
        "resourcePackages": [],
        "sections": sections,
    }, indent=2, ensure_ascii=False) + "\n")

    # --- loose DAX for manual setup ---
    dax = ["// Measures for the Shipments table. Generated by powerbi/tools/build_pbip.py.", ""]
    for m, expr, fmt, folder in MEASURES:
        dax += [f"// [{folder}] format: {fmt}", f"{m} = {expr}", ""]
    write(ROOT / "dax" / "measures.dax", "\n".join(dax))
    print("PBIP written to", ROOT)


if __name__ == "__main__":
    main()
