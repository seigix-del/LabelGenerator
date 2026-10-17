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
    ("G Total", "SUM(Shipments[GQty]) + 0", "#,0", "Grade stock"),
    ("G IN", "CALCULATE(SUM(Shipments[GQty]), Shipments[InWarehouse] = TRUE()) + 0", "#,0", "Grade stock"),
    ("G OUT", "CALCULATE(SUM(Shipments[GQty]), Shipments[InWarehouse] = FALSE()) + 0", "#,0", "Grade stock"),
    ("G Remaining", "[G Total] - [G OUT]", "#,0", "Grade stock"),
    ("L Total", "SUM(Shipments[LossQty]) + 0", "#,0", "Grade stock"),
    ("L IN", "CALCULATE(SUM(Shipments[LossQty]), Shipments[InWarehouse] = TRUE()) + 0", "#,0", "Grade stock"),
    ("L OUT", "CALCULATE(SUM(Shipments[LossQty]), Shipments[InWarehouse] = FALSE()) + 0", "#,0", "Grade stock"),
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


# ---------------------------------------------------------------------------
# Report visuals (legacy report.json layout)
# ---------------------------------------------------------------------------
GREEN, RED, NAVY = "#00B050", "#FF0000", "#002060"
GL_MEASURES = ["G Total", "G IN", "G OUT", "G Remaining", "L Total", "L IN", "L OUT", "L Remaining"]


def M(name):
    return ("Shipments", name, "Measure")


def C(table, col):
    return (table, col, "Column")


def lit(v):
    return {"expr": {"Literal": {"Value": v}}}


SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition"


class Visual:
    """One visual: type, position, and fields per role ({role: [field, ...]})."""

    def __init__(self, vtype, x, y, w, h, roles=None, title=None, color=None, text=None):
        self.vtype, self.pos, self.roles = vtype, (x, y, w, h), roles or {}
        self.title, self.color, self.text = title, color, text

    def body(self):
        if self.vtype == "textbox":
            run = {"value": self.text, "textStyle": {"fontWeight": "bold", "fontSize": "20pt", "color": NAVY}}
            return {"visualType": "textbox", "drillFilterOtherVisuals": True,
                    "objects": {"general": [{"properties": {"paragraphs": [{"textRuns": [run]}]}}]}}
        state = {}
        for role, fields in self.roles.items():
            projections = []
            for table, prop, kind in fields:
                proj = {"field": {kind: {"Expression": {"SourceRef": {"Entity": table}}, "Property": prop}},
                        "queryRef": f"{table}.{prop}", "nativeQueryRef": prop}
                if role in ("Rows", "Columns", "Category", "Series") or self.vtype == "slicer":
                    proj["active"] = True
                projections.append(proj)
            state[role] = {"projections": projections}
        v = {"visualType": self.vtype, "query": {"queryState": state}, "drillFilterOtherVisuals": True}
        if self.color and self.vtype == "card":
            v["objects"] = {"labels": [{"properties": {"color": {"solid": {"color": lit(f"'{self.color}'")}}}}]}
        if self.title:
            v["visualContainerObjects"] = {"title": [{"properties": {"show": lit("true"),
                                                                     "text": lit(f"'{self.title}'")}}]}
        return v

    def visual_json(self, page, z):
        x, y, w, h = self.pos
        name = uuid.uuid5(NS, f"visual/{page}/{z}").hex[:20]
        return name, {"$schema": f"{SCHEMA}/visualContainer/1.6.0/schema.json", "name": name,
                      "position": {"x": x, "y": y, "z": z * 1000, "height": h, "width": w, "tabOrder": z * 1000},
                      "visual": self.body()}


def card(measure, x, y, color=None):
    return Visual("card", x, y, 200, 110, {"Values": [M(measure)]}, title=measure, color=color)


def slicer(field, x, y, w=300, h=150, title=None):
    return Visual("slicer", x, y, w, h, {"Values": [field]}, title=title)


def title_box(text):
    return Visual("textbox", 10, 5, 1260, 50, text=text)


def report_pages():
    grade_matrix_values = [M(m) for m in GL_MEASURES]
    shipment_cols = [C("Shipments", c) for c in ("Waybill", "Reference", "SecondaryRef", "Contents", "Grade", "RowType",
                                                  "LastEventLabel", "Location", "City", "Consignee", "StatusFileDate")]
    return [
        ("Dashboard", [
            title_box("Inventory dashboard"),
            card("IN Waybills", 10, 60, GREEN), card("OUT Waybills", 220, 60, RED),
            card("Total Waybills", 430, 60), card("% IN", 640, 60),
            card("Unmatched Waybills", 850, 60), card("Last Status File Date", 1060, 60),
            Visual("clusteredBarChart", 10, 185, 500, 525,
                   {"Category": [C("Shipments", "LastEventLabel")], "Series": [C("Shipments", "Location")],
                    "Y": [M("Total Waybills")]}, title="Waybills by last event"),
            Visual("pivotTable", 520, 185, 750, 370,
                   {"Rows": [C("Grades", "Grade")], "Values": grade_matrix_values}, title="Grade stock"),
            slicer(C("Shipments", "Consignee"), 520, 565, 370, 145, "Consignee"),
            slicer(C("Shipments", "City"), 900, 565, 370, 145, "City"),
        ]),
        ("Grade stock", [
            title_box("Grade stock - G and L (Remaining = Total - OUT)"),
            Visual("pivotTable", 10, 60, 1260, 520,
                   {"Rows": [C("Grades", "Grade")], "Values": grade_matrix_values}, title="Grade stock"),
            card("No Grade Waybills", 10, 595), card("Total Waybills", 220, 595),
            slicer(C("Shipments", "Consignee"), 430, 595, 420, 115, "Consignee"),
            slicer(C("Shipments", "City"), 860, 595, 410, 115, "City"),
        ]),
        ("Consignee x grade", [
            title_box("Consignee x grade (G Total and L Total per grade)"),
            Visual("pivotTable", 10, 60, 1260, 520,
                   {"Rows": [C("Shipments", "Consignee"), C("Shipments", "CustomerCode")],
                    "Columns": [C("Grades", "Grade")], "Values": [M("G Total"), M("L Total")]},
                   title="Consignee x grade"),
            slicer(C("Shipments", "Location"), 10, 595, 300, 115, "Location"),
            slicer(C("Shipments", "City"), 320, 595, 400, 115, "City"),
        ]),
        ("Shipments", [
            title_box("Shipments"),
            slicer(C("Shipments", "Location"), 10, 60, 250, 120, "Location"),
            slicer(C("Grades", "Grade"), 270, 60, 250, 120, "Grade"),
            slicer(C("Shipments", "Consignee"), 530, 60, 470, 120, "Consignee"),
            card("Total Waybills", 1010, 65),
            Visual("tableEx", 10, 190, 1260, 520, {"Values": shipment_cols}, title="Waybills"),
        ]),
        ("Unmatched", [
            title_box("Unmatched - in a status file but not in the inventory list"),
            card("Unmatched Waybills", 10, 60, RED),
            Visual("tableEx", 10, 185, 1260, 525,
                   {"Values": [C("Unmatched", c) for c in ("Waybill", "LastEvent", "CityName", "ConsigneeName",
                                                           "SourceFile", "FileDate")]}, title="Unmatched waybills"),
        ]),
    ]


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
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
        "version": "1.0",
        "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
        "settings": {"enableAutoRecovery": True},
    }, indent=2) + "\n")

    # --- semantic model (TMDL) ---
    write(sm / "definition.pbism", json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
        "version": "4.2", "settings": {}}, indent=2) + "\n")
    write(sm / "definition" / "database.tmdl", "database\n\tcompatibilityLevel: 1601\n")
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

    # --- report (PBIR folder layout, as saved by Power BI Desktop) ---
    write(rp / "definition.pbir", json.dumps({
        "$schema": f"{SCHEMA}Properties/1.0.0/schema.json",
        "version": "4.0",
        "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}},
    }, indent=2) + "\n")
    d = rp / "definition"
    write(d / "version.json", json.dumps({"$schema": f"{SCHEMA}/versionMetadata/1.0.0/schema.json",
                                         "version": "2.0.0"}, indent=2) + "\n")
    write(d / "report.json", json.dumps({
        "$schema": f"{SCHEMA}/report/1.2.0/schema.json",
        "themeCollection": {"baseTheme": {"name": "CY24SU10", "reportVersionAtImport": "5.61",
                                          "type": "SharedResources"}},
        "layoutOptimization": "None",
        "resourcePackages": [{"name": "SharedResources", "type": "SharedResources",
                              "items": [{"name": "CY24SU10", "path": "BaseThemes/CY24SU10.json",
                                         "type": "BaseTheme"}]}],
        "settings": {"useStylableVisualContainerHeader": True, "defaultDrillFilterOtherVisuals": True,
                     "allowChangeFilterTypes": True, "useEnhancedTooltips": True},
    }, indent=2) + "\n")
    theme = rp / "StaticResources" / "SharedResources" / "BaseThemes" / "CY24SU10.json"
    theme.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(pathlib.Path(__file__).with_name("CY24SU10.json"), theme)
    page_names = []
    for i, (title, visuals) in enumerate(report_pages()):
        pname = uuid.uuid5(NS, f"page/{i}").hex[:20]
        page_names.append(pname)
        pd = d / "pages" / pname
        write(pd / "page.json", json.dumps({
            "$schema": f"{SCHEMA}/page/1.3.0/schema.json", "name": pname, "displayName": title,
            "displayOption": "FitToPage", "height": 720, "width": 1280,
        }, indent=2, ensure_ascii=False) + "\n")
        for z, v in enumerate(visuals):
            vname, body = v.visual_json(i, z)
            write(pd / "visuals" / vname / "visual.json", json.dumps(body, indent=2, ensure_ascii=False) + "\n")
    write(d / "pages" / "pages.json", json.dumps({
        "$schema": f"{SCHEMA}/pagesMetadata/1.0.0/schema.json",
        "pageOrder": page_names, "activePageName": page_names[0],
    }, indent=2) + "\n")

    # --- loose DAX for manual setup ---
    dax = ["// Measures for the Shipments table. Generated by powerbi/tools/build_pbip.py.", ""]
    for m, expr, fmt, folder in MEASURES:
        dax += [f"// [{folder}] format: {fmt}", f"{m} = {expr}", ""]
    write(ROOT / "dax" / "measures.dax", "\n".join(dax))
    print("PBIP written to", ROOT)


if __name__ == "__main__":
    main()
