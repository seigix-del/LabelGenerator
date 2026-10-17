# Inventory Tracker for Power BI

This is the Power BI version of the Access inventory tracker in [`../access`](../access). The rules are the same:

- **IN warehouse** means the last event (`آخر حدث`) is `جاهزة للالتقاط` or `استلام من العميل`. **Every other event is OUT.** A waybill that has not appeared in any status file yet counts as IN.
- **Grades** are G1–G9 and G11–G13 (there is no G10). The grade is read from `وصف محتويات الشحنة`, after `\`, `/` or `//`.
  - A plain `G2` row counts as **1** on the **G** side.
  - A `G2-L4` row counts as **4** on the **L** side, and only there.
  - G and L each have their own Total, IN, OUT and Remaining (Remaining = Total − OUT). Neither is subtracted from the other.
- The consignee is `اسم المرسل إليه`, taken from the status file.

## What's different from Access

In Power BI there is no upload button. **You drop Excel files into two folders and press Refresh.**

| Folder (parameter) | Put here | Columns read |
|---|---|---|
| `MasterFolder` | Inventory list files | `بوليصة الشحن`, `رقم المرجع`, `وصف محتويات الشحنة`, `رقم المرجع الثانوي` |
| `StatusFolder` | Status update files | `رقم التتبع`, `مدينة الوجهة`, `آخر حدث`, `اسم المرسل إليه` |

- You can keep every file in the folders. When a waybill appears in several files, **the newest file wins** (by file modified date), and within one file the last row wins. This replaces the database's "last import wins".
- The header row may be anywhere in the first 15 rows, the sheet can be any sheet, and column order doesn't matter. Headers are matched loosely, the same way as in Access.
- `StatusHistory` keeps every status row from every file, so you can see how a waybill moved over time.

## Open it

You need **Power BI Desktop** (2024 or newer, Windows). The project uses the same file layout that Desktop itself saves: TMDL for the model and PBIR for the report. Each page is a folder under `InventoryTracker.Report/definition/pages`, with one `visual.json` per visual.

1. Copy the whole `powerbi` folder to your PC.
2. **Extract the zip first** (right-click → Extract All). Don't open the file from inside the zip. If Desktop says it can't open the project, go to **File → Options → Preview features**, tick **Power BI Project (.pbip) save option**, **Store semantic model using TMDL format** and **Store reports using enhanced metadata format (PBIR)**, then restart Desktop. Recent versions have these on already.
3. In Desktop, **File → Open → Browse**, choose `InventoryTracker.pbip`. You should see five page tabs: Dashboard, Grade stock, Consignee x grade, Shipments and Unmatched.
4. Go to **Home → Transform data → Edit parameters**. Set `MasterFolder` and `StatusFolder` to your two folders. To try the samples, point them at `powerbi\samples\Master` and `powerbi\samples\Status`.
5. Click **Refresh**.
6. **File → Save as → .pbix** if you want a single file.

**If the project will not open.** The files were checked against Microsoft's published PBIR JSON schemas, and the model loads in Microsoft's own TMDL parser (the Tabular Object Model library). It still hasn't been opened in Power BI Desktop itself. If that happens, build the model by hand. It takes about 10 minutes:
1. In a new report, open **Transform data**.
2. For each file in [`queries/`](queries), create a **Blank query**, open the **Advanced editor**, paste the file and name the query exactly like the file.
3. Mark `MasterFolder` and `StatusFolder` as parameters. Untick **Enable load** on everything except `Shipments`, `Unmatched`, `StatusHistory` and `Grades`.
4. **Close & Apply**.
5. Relate `Shipments[Grade]` (many) to `Grades[Grade]` (one), and set `Grades[Grade]` to **Sort by column → SortOrder**.
6. Add the measures from [`dax/measures.dax`](dax/measures.dax) to `Shipments`.

## Model

```
Grades (1) ──< Shipments (one row per waybill in the inventory list)
Unmatched       (status rows whose waybill is not in the inventory list)
StatusHistory   (every status row of every file)
```

| `Shipments` column | Meaning |
|---|---|
| Waybill, Reference, SecondaryRef, Contents | From the inventory list (`Reference` defaults to the waybill) |
| Grade | G1–G9 or G11–G13, blank when there is no valid code |
| RowType | `G` (plain row) or `L` (`-L<k>` loss row) |
| GQty / LossQty | 1 / 0 for a G row, 0 / k for an L row (hidden; the measures use them) |
| LastEventLabel, HasStatus, InWarehouse, Location | Latest status and the IN/OUT rule (`(no status yet)` = IN) |
| City, Consignee, CustomerCode | From the latest status (`M174562` is parsed from the consignee name) |
| MasterFile / StatusFile (+ dates) | The file each value came from |

| Measures | |
|---|---|
| **Inventory** | `Total Waybills`, `IN Waybills`, `OUT Waybills`, `% IN`, `Unmatched Waybills`, `No Grade Waybills`, `Last Status File Date` |
| **Grade stock** | `G Total`, `G IN`, `G OUT`, `G Remaining`, `L Total`, `L IN`, `L OUT`, `L Remaining` |

## Report pages

The project already contains these pages and visuals:

| Page | Visuals |
|---|---|
| **Dashboard** | Cards: IN, OUT, Total, % IN, Unmatched, last status file date. Bar chart of waybills by last event (IN/OUT). Grade stock matrix. Consignee and City slicers. |
| **Grade stock** | Matrix with a row per grade and the TOTAL row: G Total, G IN, G OUT, G Remaining, L Total, L IN, L OUT, L Remaining. Cards for no-grade waybills and total waybills. Consignee and City slicers. |
| **Consignee x grade** | Matrix with consignees (and their code) as rows, the grades as columns, and G Total and L Total as values. Location and City slicers. |
| **Shipments** | Every waybill with grade, row type, last event, IN/OUT, city, consignee and status date. Location, Grade and Consignee slicers. |
| **Unmatched** | Waybills that are in a status file but not in the inventory list. |

The G and L measures return 0 rather than blank, so every grade (G1–G9, G11–G13) is listed even before it has any items. You can restyle or move anything in Desktop.

**The pages show nothing until the data loads.** Set both folder parameters (step 4 above) and click **Refresh**. If a visual shows an error, open **Transform data**: the query with a yellow warning shows the reason, usually a folder that doesn't exist.

## Check with the sample files

With the parameters pointing at `samples\Master` and `samples\Status`, you should get these numbers. They match the Access version.

- Total 9 waybills: **IN 4, OUT 5**. **1 unmatched** (`EBUXSL2079999999`).
- Grade stock:

  | Grade | G Total | G IN | G OUT | G Remaining | L Total | L IN | L OUT | L Remaining |
  |---|---|---|---|---|---|---|---|---|
  | G1 | 5 | 2 | 3 | 2 | 4 | 0 | 4 | 0 |
  | G9 | 2 | 2 | 0 | 2 | 1 | 0 | 1 | 0 |

## Editing

The `.pq` files in `queries/` are the source. After you change one, run `python powerbi/tools/build_pbip.py`. It regenerates the semantic model and `dax/measures.dax`, and it **recreates the report folder**. So once you've changed the visuals in Desktop, change queries in Desktop instead of re-running the script.
