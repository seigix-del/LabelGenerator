// Renderer: state, file loading/persistence, and the six pages.
(function () {
  "use strict";
  const I = window.Inventory;

  const S = {
    files: [],          // [{name, lastModified, kind, sheet, rows, sample?}]
    sample: false,
    view: "dashboard",
    filter: { consignee: "", city: "" },
    ship: { q: "", loc: "all", grade: "", limit: 300 },
    metric: "gTotal",
    confirmClear: false,
    data: null,         // cached compute() result
  };

  const TITLES = {
    dashboard: "Dashboard", grades: "Grade stock", consignees: "Consignee × grade",
    shipments: "Shipments", unmatched: "Unmatched", files: "Files",
  };
  const METRICS = [
    ["gTotal", "G Total"], ["gRem", "G Remaining"], ["lTotal", "L Total"], ["lRem", "L Remaining"],
  ];

  const $ = (id) => document.getElementById(id);
  const esc = (v) => String(v ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const fmt = (n) => Number(n || 0).toLocaleString("en-US");
  const fmtDate = (t) => (t ? new Date(t).toLocaleString("en-GB", { dateStyle: "medium", timeStyle: "short" }) : "—");
  const numCell = (n, cls = "") => `<td class="num ${n ? cls : "zero"}">${fmt(n)}</td>`;

  // ---------------------------------------------------------------- storage
  const DB = {
    db: null,
    async open() {
      if (this.db) return this.db;
      this.db = await new Promise((resolve, reject) => {
        const req = indexedDB.open("inventory-tracker", 1);
        req.onupgradeneeded = () => req.result.createObjectStore("files", { keyPath: "name" });
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
      return this.db;
    },
    async tx(mode, fn) {
      try {
        const db = await this.open();
        return await new Promise((resolve, reject) => {
          const t = db.transaction("files", mode);
          const out = fn(t.objectStore("files"));
          t.oncomplete = () => resolve(out && out.result !== undefined ? out.result : undefined);
          t.onerror = () => reject(t.error);
        });
      } catch (e) {
        console.warn("Storage unavailable", e);
        return undefined;
      }
    },
    all() { return this.tx("readonly", (s) => s.getAll()); },
    put(f) { return this.tx("readwrite", (s) => s.put(f)); },
    del(name) { return this.tx("readwrite", (s) => s.delete(name)); },
    clear() { return this.tx("readwrite", (s) => s.clear()); },
  };

  function sampleFiles() {
    return window.SAMPLE_FILES.map((f) => {
      const r = I.readSheet(f.aoa);
      return { name: f.name, lastModified: f.lastModified, kind: r.kind, sheet: "Sheet1", rows: r.rows, sample: true };
    });
  }

  function setFiles(files, sample) {
    S.files = files;
    S.sample = sample;
    S.data = null;
  }

  function data() {
    if (!S.data) {
      const { shipments, unmatched } = I.compute(S.files);
      S.data = { shipments, unmatched };
    }
    const { shipments, unmatched } = S.data;
    const f = S.filter;
    const filtered = shipments.filter((s) => (!f.consignee || s.consignee === f.consignee) && (!f.city || s.city === f.city));
    return { all: shipments, shipments: filtered, unmatched };
  }

  // ---------------------------------------------------------------- uploads
  async function addFiles(list) {
    const added = [], errors = [];
    for (const file of list) {
      if (file.name.startsWith("~$")) continue;
      if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) { errors.push(`${file.name}: not an Excel file.`); continue; }
      try {
        const res = I.parseWorkbook(XLSX, new Uint8Array(await file.arrayBuffer()));
        if (res.error) { errors.push(`${file.name}: ${res.error}`); continue; }
        added.push({ name: file.name, lastModified: file.lastModified, kind: res.kind, sheet: res.sheet, rows: res.rows });
      } catch (e) {
        errors.push(`${file.name}: could not be read (${e.message}).`);
      }
    }
    if (added.length) {
      let files = S.sample ? [] : S.files.slice();
      for (const f of added) files = files.filter((x) => x.name !== f.name).concat(f);
      setFiles(files, false);
      for (const f of added) await DB.put(f);
    }
    const lines = added.map((f) => `${f.name}: ${f.kind === "master" ? "inventory list" : "status update"}, ${fmt(f.rows.length)} rows`);
    toast([added.length ? `Added ${added.length} file${added.length > 1 ? "s" : ""}.` : "", ...lines, ...errors].filter(Boolean).join("\n"), errors.length ? 9000 : 4500);
    render();
  }

  window.openFilePicker = () => $("file-input").click();

  // ---------------------------------------------------------------- toolbar
  function filterControls(d) {
    const consignees = [...new Set(d.all.map((s) => s.consignee))].sort((a, b) => a.localeCompare(b, "ar"));
    const cities = [...new Set(d.all.map((s) => s.city).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ar"));
    const opt = (v, sel) => `<option value="${esc(v)}"${v === sel ? " selected" : ""}>${esc(v)}</option>`;
    return `
      <label class="field" for="f-consignee">Consignee
        <select id="f-consignee" dir="auto"><option value="">All</option>${consignees.map((c) => opt(c, S.filter.consignee)).join("")}</select></label>
      <label class="field" for="f-city">City
        <select id="f-city" dir="auto"><option value="">All</option>${cities.map((c) => opt(c, S.filter.city)).join("")}</select></label>`;
  }

  function tableTools() {
    return `<button type="button" class="btn" data-action="copy">Copy table</button>
            <button type="button" class="btn" data-action="export">Export to Excel</button>`;
  }

  // ---------------------------------------------------------------- pages
  function kpi(label, value, note, cls = "", extra = "", valueCls = "") {
    return `<div class="kpi ${cls}"><div class="label">${label}</div><div class="value ${valueCls}">${value}</div><div class="note">${note}</div>${extra}</div>`;
  }

  function viewDashboard(d) {
    const sm = I.summary(d.shipments, d.unmatched);
    const ev = I.byEvent(d.shipments);
    const max = Math.max(1, ...ev.map((e) => e.count));
    const gs = I.gradeStock(d.shipments);
    const pct = Math.round(sm.inPct * 1000) / 10;
    return `
      <div class="kpis">
        ${kpi("IN warehouse", fmt(sm.inCount), "Ready for pickup, received from customer, or no status yet", "in")}
        ${kpi("OUT", fmt(sm.outCount), "Any other last event", "out")}
        ${kpi("Total waybills", fmt(sm.total), "On the inventory list")}
        ${kpi("% IN", `${pct}%`, `${fmt(sm.inCount)} of ${fmt(sm.total)}`, "", `<div class="split"><div class="a" style="width:${pct}%"></div><div class="b" style="width:${100 - pct}%"></div></div>`)}
        ${kpi("Unmatched", fmt(sm.unmatched), "In a status file, not in the inventory list", sm.unmatched ? "warn" : "")}
        ${kpi("Last status update", sm.lastStatusDate ? new Date(sm.lastStatusDate).toLocaleDateString("en-GB", { dateStyle: "medium" }) : "—", "Newest status file used", "", "", "date")}
      </div>
      <div class="grid-2">
        <div class="panel">
          <div class="panel-head"><h2>Waybills by last event</h2><span class="sub">آخر حدث</span></div>
          <div class="panel-body bars">
            ${ev.length ? ev.map((e) => `
              <div class="bar-row">
                <div>
                  <div class="bar-label"><span dir="auto">${esc(e.event)}</span><span class="pill ${e.inWarehouse ? "in" : "out"}">${e.inWarehouse ? "IN" : "OUT"}</span></div>
                  <div class="bar-track"><div class="bar-fill ${e.inWarehouse ? "in" : "out"}" style="width:${(e.count / max) * 100}%"></div></div>
                </div>
                <div class="bar-num">${fmt(e.count)}</div>
              </div>`).join("") : `<div class="empty">No waybills match the filters.</div>`}
          </div>
        </div>
        <div class="panel">
          <div class="panel-head"><h2>Grade stock</h2><button type="button" class="btn small" data-view="grades">Open grade stock →</button></div>
          <div class="panel-body table-wrap">${gradeTable(gs, true)}</div>
        </div>
      </div>`;
  }

  function gradeTable(gs, compact) {
    const cols = compact
      ? [["gTotal", "G Total"], ["gOut", "G OUT", "out-num"], ["gRem", "G Remaining", "rem"], ["lTotal", "L Total", "", true], ["lOut", "L OUT", "out-num", true], ["lRem", "L Remaining", "rem", true]]
      : [["gTotal", "Total"], ["gIn", "IN", "in-num"], ["gOut", "OUT", "out-num"], ["gRem", "Remaining", "rem"],
         ["lTotal", "Total", "", true], ["lIn", "IN", "in-num", true], ["lOut", "OUT", "out-num", true], ["lRem", "Remaining", "rem", true]];
    const cell = (r, [k, , cls = "", l]) => `<td class="num ${r[k] ? cls : "zero"} ${l ? "lcol" : ""}">${fmt(r[k])}</td>`;
    const groupRow = compact ? "" : `<tr><th></th><th class="group g" colspan="4">G — plain rows, 1 each</th><th class="group l" colspan="4">L — loss rows (G‑L<i>n</i>), n each</th></tr>`;
    return `
      <table data-export="Grade stock">
        <thead>${groupRow}<tr><th>Grade</th>${cols.map(([, h]) => `<th class="num">${h}</th>`).join("")}</tr></thead>
        <tbody>${gs.rows.map((r) => `<tr><td><span class="chip">${r.grade}</span></td>${cols.map((c) => cell(r, c)).join("")}</tr>`).join("")}</tbody>
        <tfoot><tr><td>TOTAL</td>${cols.map((c) => cell(gs.total, c)).join("")}</tr></tfoot>
      </table>`;
  }

  function viewGrades(d) {
    const gs = I.gradeStock(d.shipments);
    return `
      <div class="panel">
        <div class="panel-head"><h2>Grade stock</h2><span class="sub">Remaining = Total − OUT, separately for G and L</span></div>
        <div class="panel-body table-wrap">${gradeTable(gs, false)}
          <p class="foot-note">${fmt(gs.noGrade)} waybill${gs.noGrade === 1 ? "" : "s"} without a valid grade code (G1–G9, G11–G13) ${gs.noGrade === 1 ? "is" : "are"} not counted here.</p>
        </div>
      </div>`;
  }

  function viewConsignees(d) {
    const mx = I.consigneeMatrix(d.shipments);
    const k = S.metric;
    const seg = `<div class="seg" role="group" aria-label="Value">${METRICS.map(([m, label]) =>
      `<button type="button" data-metric="${m}" aria-pressed="${m === k}">${label}</button>`).join("")}</div>`;
    const label = METRICS.find(([m]) => m === k)[1];
    return `
      <div class="panel">
        <div class="panel-head"><h2>${label} by consignee and grade</h2>${seg}</div>
        <div class="panel-body table-wrap">
          ${mx.consignees.length ? `
          <table data-export="Consignee x grade">
            <thead><tr><th>اسم المرسل إليه</th><th>Code</th>${I.GRADES.map((g) => `<th class="num">${g}</th>`).join("")}<th class="num">Total</th></tr></thead>
            <tbody>${mx.consignees.map((c) => `
              <tr><td dir="auto">${esc(c.name)}</td><td class="mono">${esc(c.code)}</td>
                ${I.GRADES.map((g) => numCell(c.cells.get(g)[k])).join("")}${numCell(c.total[k], "rem")}</tr>`).join("")}
            </tbody>
            <tfoot><tr><td>TOTAL</td><td></td>${I.GRADES.map((g) => numCell(mx.colTotals.get(g)[k])).join("")}${numCell(mx.grand[k], "rem")}</tr></tfoot>
          </table>` : `<div class="empty">No graded waybills match the filters.</div>`}
        </div>
      </div>`;
  }

  function viewShipments(d) {
    const q = I.normKey(S.ship.q);
    let rows = d.shipments.filter((s) =>
      (S.ship.loc === "all" || s.location === S.ship.loc) &&
      (!S.ship.grade || s.grade === S.ship.grade) &&
      (!q || [s.waybill, s.reference, s.secondaryRef, s.contents, s.lastEvent, s.city, s.consignee, s.customerCode]
        .some((v) => I.normKey(v).includes(q))));
    const total = rows.length;
    rows = rows.slice(0, S.ship.limit);
    const seg = `<div class="seg" role="group" aria-label="Location">${[["all", "All"], ["IN", "IN"], ["OUT", "OUT"]].map(([v, l]) =>
      `<button type="button" data-loc="${v}" aria-pressed="${S.ship.loc === v}">${l}</button>`).join("")}</div>`;
    return `
      <div class="panel">
        <div class="panel-head" style="flex-wrap:wrap">
          <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">
            <input type="search" id="ship-q" placeholder="Search waybill, reference, consignee…" value="${esc(S.ship.q)}" dir="auto">
            ${seg}
            <select id="ship-grade" aria-label="Grade"><option value="">All grades</option>${I.GRADES.map((g) =>
              `<option${g === S.ship.grade ? " selected" : ""}>${g}</option>`).join("")}</select>
          </div>
          <span class="sub">${fmt(total)} waybill${total === 1 ? "" : "s"}</span>
        </div>
        <div class="table-wrap">
          ${total ? `<table data-export="Shipments">
            <thead><tr><th>Waybill</th><th>Reference</th><th>Secondary ref</th><th>Contents</th><th>Grade</th><th>Last event</th><th>Location</th><th>City</th><th>اسم المرسل إليه</th><th>Status file date</th></tr></thead>
            <tbody>${rows.map((s) => `
              <tr><td class="mono">${esc(s.waybill)}</td><td class="mono">${esc(s.reference)}</td><td class="mono">${esc(s.secondaryRef)}</td>
                <td dir="auto" class="wrap">${esc(s.contents)}</td>
                <td>${s.grade ? `<span class="chip">${s.grade}</span>` : ""}${s.lossQty ? ` <span class="chip l">L${s.lossQty}</span>` : ""}</td>
                <td dir="auto">${esc(s.lastEvent)}</td><td><span class="pill ${s.inWarehouse ? "in" : "out"}">${s.location}</span></td>
                <td dir="auto">${esc(s.city)}</td><td dir="auto">${esc(s.consignee)}</td><td>${fmtDate(s.statusDate)}</td></tr>`).join("")}
            </tbody></table>` : `<div class="empty">No waybills match.</div>`}
        </div>
        ${total > rows.length ? `<div class="more">Showing ${fmt(rows.length)} of ${fmt(total)}. <button type="button" class="btn small" data-action="more">Show all</button></div>` : ""}
      </div>`;
  }

  function viewUnmatched(d) {
    const u = d.unmatched;
    return `
      <div class="panel">
        <div class="panel-head"><h2>In a status file but not in the inventory list</h2><span class="sub">${fmt(u.length)} waybill${u.length === 1 ? "" : "s"}</span></div>
        <div class="table-wrap">
          ${u.length ? `<table data-export="Unmatched">
            <thead><tr><th>Waybill</th><th>Last event</th><th>City</th><th>اسم المرسل إليه</th><th>Status file</th><th>File date</th></tr></thead>
            <tbody>${u.map((r) => `<tr><td class="mono">${esc(r.waybill)}</td><td dir="auto">${esc(r.lastEvent)}</td><td dir="auto">${esc(r.city)}</td>
              <td dir="auto">${esc(r.consignee)}</td><td>${esc(r.file)}</td><td>${fmtDate(r.fileDate)}</td></tr>`).join("")}</tbody>
          </table>` : `<div class="empty">Every waybill in the status files is on the inventory list.</div>`}
        </div>
      </div>`;
  }

  function viewFiles() {
    const files = S.files.slice().sort((a, b) => b.lastModified - a.lastModified);
    const clear = S.confirmClear
      ? `<span class="confirm">Remove all files? <button type="button" class="btn danger small" data-action="clear-yes">Remove all</button><button type="button" class="btn small" data-action="clear-no">Cancel</button></span>`
      : `<button type="button" class="btn small" data-action="clear">Remove all</button>`;
    return `
      <div class="dropzone">
        <strong>Drop Excel files here, or</strong>
        <button type="button" class="btn primary" data-action="upload">Choose files…</button>
        <span class="hint">Add inventory lists and status updates at any time. The type is detected from the headers.
        When a waybill is in several files, the newest file wins.</span>
      </div>
      <div class="panel">
        <div class="panel-head"><h2>${S.sample ? "Sample files (replaced by your first upload)" : "Loaded files"}</h2>
          <span class="sub">${S.sample ? `<button type="button" class="btn small" data-action="upload">Upload your files</button>` : clear}</span></div>
        <div class="table-wrap">
          ${files.length ? `<table>
            <thead><tr><th>File</th><th>Type</th><th>Sheet</th><th class="num">Rows</th><th>Modified</th><th></th></tr></thead>
            <tbody>${files.map((f) => `<tr><td>${esc(f.name)}</td>
              <td><span class="chip kind-${f.kind}">${f.kind === "master" ? "Inventory list" : "Status update"}</span></td>
              <td>${esc(f.sheet)}</td><td class="num">${fmt(f.rows.length)}</td><td>${fmtDate(f.lastModified)}</td>
              <td>${S.sample ? "" : `<button type="button" class="btn small" data-remove="${esc(f.name)}">Remove</button>`}</td></tr>`).join("")}</tbody>
          </table>` : `<div class="empty">No files yet. Upload an inventory list (بوليصة الشحن …) and a status update (رقم التتبع, آخر حدث …).
            <p><button type="button" class="btn small" data-action="sample">Load sample data</button></p></div>`}
        </div>
      </div>
      <div class="rules">
        <div class="panel panel-body"><div class="rule"><h3>IN / OUT</h3><p>IN when آخر حدث is جاهزة للالتقاط or استلام من العميل, or the waybill has no status yet. Every other event is OUT.</p></div></div>
        <div class="panel panel-body"><div class="rule"><h3>Grades</h3><p>From وصف محتويات الشحنة after \\, / or //. G1–G9 and G11–G13. A G2 row counts 1 on the G side; G2‑L4 counts 4 on the L side only. A letter prefix is ignored (MG3‑L1 = G3, loss 1).</p></div></div>
        <div class="panel panel-body"><div class="rule"><h3>Consignee</h3><p>اسم المرسل إليه from the newest status file. The code at the end (e.g. M174562) is shown separately.</p></div></div>
      </div>`;
  }

  // ---------------------------------------------------------------- render
  function render() {
    const d = data();
    document.querySelectorAll("#nav [data-view]").forEach((b) => {
      if (b.dataset.view === S.view) b.setAttribute("aria-current", "page");
      else b.removeAttribute("aria-current");
    });
    $("nav-shipments").textContent = fmt(d.all.length);
    $("nav-unmatched").textContent = d.unmatched.length ? fmt(d.unmatched.length) : "";
    $("nav-files").textContent = fmt(S.files.length);
    $("page-title").textContent = TITLES[S.view];
    $("sample-pill").hidden = !S.sample;
    $("side-hint").textContent = S.sample
      ? "Showing sample data. Upload your inventory list and status files to replace it."
      : `${fmt(S.files.filter((f) => f.kind === "master").length)} inventory, ${fmt(S.files.filter((f) => f.kind === "status").length)} status file(s) loaded.`;

    const withFilters = ["dashboard", "grades", "consignees", "shipments"].includes(S.view);
    const withTable = ["grades", "consignees", "shipments", "unmatched"].includes(S.view);
    $("toolbar-tools").innerHTML = (withFilters ? filterControls(d) : "") + (withTable ? tableTools() : "");

    const views = { dashboard: viewDashboard, grades: viewGrades, consignees: viewConsignees, shipments: viewShipments, unmatched: viewUnmatched, files: viewFiles };
    $("view").innerHTML = views[S.view](d);
  }

  function go(view) {
    S.view = view;
    S.confirmClear = false;
    render();
    $("view").scrollTop = 0;
  }

  // ---------------------------------------------------------------- copy / export
  function currentTable() {
    return $("view").querySelector("table[data-export]");
  }

  async function copyTable() {
    const t = currentTable();
    if (!t) return toast("Nothing to copy on this page.");
    const tsv = [...t.rows].map((r) => [...r.cells].map((c) => c.innerText.replace(/\s+/g, " ").trim()).join("\t")).join("\n");
    try {
      await navigator.clipboard.writeText(tsv);
      toast("Table copied. Paste it into Excel.");
    } catch {
      toast("Could not copy. Use Export to Excel instead.");
    }
  }

  function exportTable() {
    const t = currentTable();
    if (!t) return toast("Nothing to export on this page.");
    const name = t.dataset.export;
    const wb = XLSX.utils.table_to_book(t, { sheet: name.slice(0, 31), raw: true });
    XLSX.writeFile(wb, `${name} ${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  // ---------------------------------------------------------------- events
  let toastTimer;
  function toast(msg, ms = 3500) {
    const el = $("toast");
    el.textContent = msg;
    el.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { el.hidden = true; }, ms);
  }

  document.addEventListener("click", async (e) => {
    const t = e.target.closest("[data-view],[data-action],[data-metric],[data-loc],[data-remove]");
    if (!t) return;
    if (t.dataset.view) return go(t.dataset.view);
    if (t.dataset.metric) { S.metric = t.dataset.metric; return render(); }
    if (t.dataset.loc) { S.ship.loc = t.dataset.loc; S.ship.limit = 300; return render(); }
    if (t.dataset.remove) {
      setFiles(S.files.filter((f) => f.name !== t.dataset.remove), false);
      await DB.del(t.dataset.remove);
      toast(`Removed ${t.dataset.remove}.`);
      return render();
    }
    switch (t.dataset.action) {
      case "upload": return window.openFilePicker();
      case "copy": return copyTable();
      case "export": return exportTable();
      case "more": S.ship.limit = Infinity; return render();
      case "clear": S.confirmClear = true; return render();
      case "clear-no": S.confirmClear = false; return render();
      case "clear-yes":
        await DB.clear();
        setFiles([], false);
        S.confirmClear = false;
        toast("All files removed.");
        return render();
      case "sample": setFiles(sampleFiles(), true); return render();
    }
  });

  document.addEventListener("change", (e) => {
    if (e.target.id === "f-consignee") { S.filter.consignee = e.target.value; render(); }
    if (e.target.id === "f-city") { S.filter.city = e.target.value; render(); }
    if (e.target.id === "ship-grade") { S.ship.grade = e.target.value; render(); }
    if (e.target.id === "file-input") { addFiles([...e.target.files]); e.target.value = ""; }
  });

  let searchTimer;
  document.addEventListener("input", (e) => {
    if (e.target.id !== "ship-q") return;
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      S.ship.q = e.target.value;
      S.ship.limit = 300;
      const pos = e.target.selectionStart;
      render();
      const box = $("ship-q");
      box.focus();
      box.setSelectionRange(pos, pos);
    }, 200);
  });

  // Drag & drop anywhere in the window.
  let dragDepth = 0;
  window.addEventListener("dragenter", (e) => { if (e.dataTransfer?.types?.includes("Files")) { dragDepth++; $("drop-overlay").hidden = false; } });
  window.addEventListener("dragleave", () => { if (--dragDepth <= 0) { dragDepth = 0; $("drop-overlay").hidden = true; } });
  window.addEventListener("dragover", (e) => e.preventDefault());
  window.addEventListener("drop", (e) => {
    e.preventDefault();
    dragDepth = 0;
    $("drop-overlay").hidden = true;
    if (e.dataTransfer?.files?.length) addFiles([...e.dataTransfer.files]);
  });

  // ---------------------------------------------------------------- start
  (async function start() {
    const saved = (await DB.all()) || [];
    if (saved.length) setFiles(saved, false);
    else setFiles(sampleFiles(), true);
    render();
  })();
})();
