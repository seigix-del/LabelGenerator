# Inventory Tracker for macOS

A desktop app version of the inventory dashboard. You add your Excel files, and it shows IN/OUT, grade stock and consignee totals. It runs fully offline. Nothing is uploaded anywhere: files are read on your Mac, and the parsed rows are kept in the app's local storage.

It uses the same rules as the Access and Power BI versions:

- **IN** when the last event (`آخر حدث`) is `جاهزة للالتقاط` or `استلام من العميل`, or the waybill has no status yet. **Every other event is OUT.**
- **Grades** G1–G9 and G11–G13 are read from `وصف محتويات الشحنة`, after `\`, `/` or `//`. A `G2` row counts 1 on the **G** side. A `G2-L4` row counts 4 on the **L** side only. Remaining = Total − OUT, worked out separately for G and L.
- The consignee comes from `اسم المرسل إليه`.
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
- **Pages:** Dashboard, Grade stock, Consignee × grade, Shipments (search, IN/OUT and grade filters), Unmatched, and Files (remove files or start over).
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
