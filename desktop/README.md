# Inventory Tracker for macOS

A desktop app version of the inventory dashboard. You add your Excel files, and it shows IN/OUT, grade stock and consignee totals. It runs fully offline. Nothing is uploaded anywhere: files are read on your Mac, and the parsed rows are kept in the app's local storage.

It uses the same rules as the Access and Power BI versions:

- **IN** when the last event (`آخر حدث`) is `جاهزة للالتقاط` or `استلام من العميل`, or the waybill has no status yet. **Every other event is OUT.**
- **Grades** G1–G9 and G11–G13 are read from `وصف محتويات الشحنة`, after `\`, `/` or `//`. A `G2` row counts 1 on the **G** side. A `G2-L4` row counts 4 on the **L** side only. A one- or two-letter prefix is ignored, so `MG3-L1` counts as G3 with a loss of 1. Remaining = Total − OUT, worked out separately for G and L.
- Event names are matched loosely. Hidden characters, Arabic presentation forms, diacritics, hamza variants and extra spaces are ignored, and an event that contains `جاهزة للالتقاط` or `استلام من العميل` also counts as IN.
- The consignee comes from `اسم المرسل إليه`.
- **Pallet group** is `الرقم المرجعي الثانوي` (for example `MKK-SEP30CH`). For each group, the **Pallet groups** page shows its items, IN, OUT, **remaining** (items − OUT) and loss. Each group is marked **Not started** (nothing OUT), **Partial** or **Complete** (every item OUT). Click a group to list its waybills.
- **Headers** are matched by meaning, not only exact text. `رقم المرجع`, `الرقم المرجعي` and other columns containing `مرجع` count as the reference, and if the header also contains `ثانوي` it is the pallet group. The **Files** page shows which header was read for each field.
- When a waybill is in several files, the newest file wins (then the last row in that file).

## Install

1. Download the zip for your Mac. Use `InventoryTracker-mac-arm64.zip` for Apple Silicon (M1–M4). Use `InventoryTracker-mac-x64.zip` for Intel.
2. Double-click the zip and drag **Inventory Tracker.app** to **Applications**.
3. **The first time only,** right-click the app → **Open** → **Open**. The app isn't signed with an Apple Developer ID, so a normal double-click is blocked once.
   - On macOS 15 and later, that right-click → Open option may not be offered. Instead, open the app once, then go to **System Settings → Privacy & Security** → **Open Anyway**.
   - If macOS says the app "is damaged", run this once in Terminal, then open it again:
     ```
     xattr -cr "/Applications/Inventory Tracker.app"
     ```

## Use

- **Upload Excel files…** (or ⌘O, or drag files onto the window) to add inventory lists and status updates. The type is detected from the headers. Add new status files whenever you get them.
- **Pages:** Dashboard, Grade stock, Consignee × grade, Shipments (search, IN/OUT, grade and pallet filters), Pallet groups, Unmatched, and Files (remove files or start over, and see which columns were read).
- **Large projects:** tested with 300,000 waybills. All the totals are calculated in about 1.5 s, and switching pages takes well under a second. Tables show the first few hundred rows, and **Show all** lists the rest. Totals always include every row.
- **Copy table** pastes into Excel. **Export to Excel** saves the current table as `.xlsx`.
- On first launch the app shows sample data, marked **Sample data**. Your first upload replaces it.

## Build

```
cd desktop
npm install
npm test               # logic tests against access/samples/*.xlsx
npm start              # run the app (on any OS)
npm run package:mac    # dist/Inventory Tracker-darwin-{arm64,x64}/Inventory Tracker.app
```

To change the icon, edit and run `python build/make_icon.py`.

The **Build macOS app** GitHub Actions workflow builds and ad-hoc signs both apps on a Mac runner, then uploads them as a downloadable artifact. Run it from the repo's **Actions** tab, or push to `main`.

Code layout:

| File | What it does |
|---|---|
| `app/logic.js` | All the rules (parsing, IN/OUT, grades, totals). Shared with the tests. |
| `app/ui.js` | The pages, file loading and local storage |
| `app/styles.css` | Styles, with light and dark mode following macOS |
| `main.js` | The app window and menu |
