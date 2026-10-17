// Inventory logic shared by the app UI and the Node tests.
// Same rules as the Access and Power BI versions:
//  - IN warehouse = last event is "جاهزة للالتقاط" or "استلام من العميل"; anything else is OUT.
//    A waybill with no status yet counts as IN (it is on the inventory list).
//  - Grades G1-G9, G11-G13 (no G10). "G2" = 1 item on the G side; "G2-L4" = 4 items on the L side only.
//  - When a waybill appears in several files, the newest file wins, then the last row in it.
(function (root) {
  "use strict";

  const IN_EVENTS = ["جاهزة للالتقاط", "استلام من العميل"];
  const GRADES = ["G1", "G2", "G3", "G4", "G5", "G6", "G7", "G8", "G9", "G11", "G12", "G13"];
  const NO_STATUS = "(no status yet)";
  const NO_CONSIGNEE = "(no consignee yet)";

  const ALIASES = {
    Waybill: ["بوليصة الشحن", "رقم التتبع", "رقم البوليصة", "Waybill", "AWB", "Tracking Number", "Tracking No"],
    ReferenceNo: ["رقم المرجع", "Reference", "Reference No", "Ref"],
    Contents: ["وصف محتويات الشحنة", "Contents", "Description", "Shipment Contents"],
    SecondaryRef: ["رقم المرجع الثانوي", "Secondary Reference", "Secondary Ref"],
    CityName: ["مدينة الوجهة", "Destination City", "City"],
    LastEvent: ["آخر حدث", "Last Event", "Status"],
    ConsigneeName: ["اسم المرسل إليه", "Consignee", "Consignee Name", "Receiver"],
  };
  const LOGICAL = Object.keys(ALIASES);

  function normText(v) {
    if (v === null || v === undefined) return "";
    return String(v)
      .replace(/[‎‏‪-‮ـ﻿]/g, "")
      .replace(/[\r\n\t ]/g, " ")
      .replace(/ {2,}/g, " ")
      .trim();
  }

  function normKey(v) {
    return normText(v)
      .replace(/[آأإ]/g, "ا")
      .replace(/ى/g, "ي")
      .replace(/ة/g, "ه")
      .toLowerCase();
  }

  function normWaybill(v) {
    return normText(v).toUpperCase().replace(/ /g, "");
  }

  const HEADER_MAP = new Map();
  for (const logical of LOGICAL) for (const h of ALIASES[logical]) HEADER_MAP.set(normKey(h), logical);
  const IN_KEYS = new Set(IN_EVENTS.map(normKey));

  function headerLogical(cell) {
    return HEADER_MAP.get(normKey(cell)) || null;
  }

  // Last grade code in the text: "… \\ G1", "… / G1", "… // G1-L4", "ASR-SEP30-G9".
  function parseGradeLoss(text) {
    const s = normText(text);
    if (!s) return null;
    const re = /\b(G\d{1,2})(?:\s*-\s*L\s*(\d{1,5}))?\b/gi;
    let m, last = null;
    while ((m = re.exec(s)) !== null) last = m;
    if (!last) return null;
    const grade = last[1].toUpperCase();
    if (!GRADES.includes(grade)) return null;
    return { grade, loss: last[2] ? parseInt(last[2], 10) : 0 };
  }

  // "<school> - مسارات - M174562" -> "M174562"
  function customerCode(name) {
    const parts = normText(name).split("-");
    if (parts.length < 2) return "";
    const tail = parts[parts.length - 1].trim();
    return /^[A-Za-z][A-Za-z0-9]*\d$/.test(tail) && tail.length <= 30 ? tail.toUpperCase() : "";
  }

  // aoa = rows of cells. Finds the header row (first 15 rows), returns logical rows.
  function readSheet(aoa) {
    const top = Math.min(aoa.length, 15);
    let hr = -1;
    for (let r = 0; r < top && hr < 0; r++) {
      if ((aoa[r] || []).some((c) => headerLogical(c) === "Waybill")) hr = r;
    }
    if (hr < 0) return null;
    const colOf = {};
    (aoa[hr] || []).forEach((c, i) => {
      const l = headerLogical(c);
      if (l && !(l in colOf)) colOf[l] = i;
    });
    const rows = [];
    for (let r = hr + 1; r < aoa.length; r++) {
      const src = aoa[r] || [];
      const row = { rowIndex: r + 1 };
      for (const l of LOGICAL) row[l] = l in colOf ? normText(src[colOf[l]]) : "";
      if (Object.values(row).some((v) => typeof v === "string" && v)) rows.push(row);
    }
    let kind = null;
    if ("LastEvent" in colOf) kind = "status";
    else if ("ReferenceNo" in colOf || "Contents" in colOf) kind = "master";
    return { kind, columns: Object.keys(colOf), rows };
  }

  // Parses a workbook with SheetJS. Returns {kind, sheet, rows} or {error}.
  function parseWorkbook(XLSX, data) {
    const wb = XLSX.read(data, { type: "array" });
    for (const name of wb.SheetNames) {
      const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: false, defval: null, blankrows: false });
      const res = readSheet(aoa);
      if (!res) continue;
      if (!res.kind) {
        return { error: `Sheet "${name}" has a waybill column but no "آخر حدث" (status file) and no "رقم المرجع" / "وصف محتويات الشحنة" (inventory file).` };
      }
      return { kind: res.kind, sheet: name, rows: res.rows };
    }
    return { error: 'No sheet has a waybill header ("بوليصة الشحن" or "رقم التتبع") in its first 15 rows.' };
  }

  // files: [{name, lastModified, kind, rows}] -> {shipments, unmatched}
  function compute(files) {
    const ordered = files
      .map((f, i) => ({ f, i }))
      .sort((a, b) => (a.f.lastModified - b.f.lastModified) || (a.i - b.i))
      .map((x) => x.f);

    const master = new Map();
    const status = new Map();
    for (const f of ordered) {
      for (const r of f.rows) {
        const wb = normWaybill(r.Waybill);
        if (!wb) continue;
        if (f.kind === "master") {
          master.delete(wb);
          master.set(wb, { ...r, file: f.name });
        } else if (f.kind === "status" && normText(r.LastEvent)) {
          status.delete(wb);
          status.set(wb, { ...r, file: f.name, fileDate: f.lastModified });
        }
      }
    }

    const shipments = [];
    for (const [wb, m] of master) {
      const gl = parseGradeLoss(m.Contents) || parseGradeLoss(m.SecondaryRef);
      const s = status.get(wb);
      const loss = gl ? gl.loss : 0;
      const inW = s ? IN_KEYS.has(normKey(s.LastEvent)) : true;
      shipments.push({
        waybill: wb,
        reference: m.ReferenceNo || wb,
        secondaryRef: m.SecondaryRef,
        contents: m.Contents,
        grade: gl ? gl.grade : "",
        rowType: loss > 0 ? "L" : "G",
        gQty: loss > 0 ? 0 : 1,
        lossQty: loss,
        lastEvent: s ? s.LastEvent : NO_STATUS,
        hasStatus: !!s,
        inWarehouse: inW,
        location: inW ? "IN" : "OUT",
        city: s ? s.CityName : "",
        consignee: s && s.ConsigneeName ? s.ConsigneeName : NO_CONSIGNEE,
        customerCode: s ? customerCode(s.ConsigneeName) : "",
        masterFile: m.file,
        statusFile: s ? s.file : "",
        statusDate: s ? s.fileDate : null,
      });
    }
    const unmatched = [];
    for (const [wb, s] of status) {
      if (!master.has(wb)) {
        unmatched.push({ waybill: wb, lastEvent: s.LastEvent, city: s.CityName, consignee: s.ConsigneeName, file: s.file, fileDate: s.fileDate });
      }
    }
    return { shipments, unmatched };
  }

  function emptyStock() {
    return { gTotal: 0, gIn: 0, gOut: 0, gRem: 0, lTotal: 0, lIn: 0, lOut: 0, lRem: 0 };
  }

  function addTo(acc, s) {
    acc.gTotal += s.gQty;
    acc.lTotal += s.lossQty;
    if (s.inWarehouse) { acc.gIn += s.gQty; acc.lIn += s.lossQty; } else { acc.gOut += s.gQty; acc.lOut += s.lossQty; }
    acc.gRem = acc.gTotal - acc.gOut;
    acc.lRem = acc.lTotal - acc.lOut;
  }

  // One row per grade (all 12 shown), plus TOTAL and the count of waybills without a valid grade.
  function gradeStock(shipments) {
    const rows = new Map(GRADES.map((g) => [g, emptyStock()]));
    const total = emptyStock();
    let noGrade = 0;
    for (const s of shipments) {
      if (!s.grade) { noGrade++; continue; }
      addTo(rows.get(s.grade), s);
      addTo(total, s);
    }
    return { rows: GRADES.map((g) => ({ grade: g, ...rows.get(g) })), total, noGrade };
  }

  function summary(shipments, unmatched) {
    const total = shipments.length;
    const inCount = shipments.filter((s) => s.inWarehouse).length;
    const dates = shipments.map((s) => s.statusDate).filter(Boolean);
    return {
      total, inCount, outCount: total - inCount,
      inPct: total ? inCount / total : 0,
      unmatched: unmatched.length,
      lastStatusDate: dates.length ? Math.max(...dates) : null,
    };
  }

  function byEvent(shipments) {
    const m = new Map();
    for (const s of shipments) {
      const e = m.get(s.lastEvent) || { event: s.lastEvent, count: 0, inWarehouse: s.inWarehouse };
      e.count++;
      m.set(s.lastEvent, e);
    }
    return [...m.values()].sort((a, b) => b.count - a.count);
  }

  // Consignee x grade: {consignees:[{name, code, cells:{G1:{...stock}}, total}], columnTotals}
  function consigneeMatrix(shipments) {
    const m = new Map();
    const colTotals = new Map(GRADES.map((g) => [g, emptyStock()]));
    const grand = emptyStock();
    for (const s of shipments) {
      if (!s.grade) continue;
      let c = m.get(s.consignee);
      if (!c) {
        c = { name: s.consignee, code: s.customerCode, cells: new Map(GRADES.map((g) => [g, emptyStock()])), total: emptyStock() };
        m.set(s.consignee, c);
      }
      addTo(c.cells.get(s.grade), s);
      addTo(c.total, s);
      addTo(colTotals.get(s.grade), s);
      addTo(grand, s);
    }
    const consignees = [...m.values()].sort((a, b) => a.name.localeCompare(b.name, "ar"));
    return { consignees, colTotals, grand };
  }

  const api = { IN_EVENTS, GRADES, NO_STATUS, NO_CONSIGNEE, normText, normKey, normWaybill, headerLogical,
    parseGradeLoss, customerCode, readSheet, parseWorkbook, compute, gradeStock, summary, byEvent, consigneeMatrix };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Inventory = api;
})(typeof window !== "undefined" ? window : globalThis);
