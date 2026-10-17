// Inventory logic shared by the app UI and the Node tests.
// Same rules as the Access and Power BI versions:
//  - IN warehouse = last event is "جاهزة للالتقاط" or "استلام من العميل"; anything else is OUT.
//    A waybill with no status yet counts as IN (it is on the inventory list).
//  - Grades G1-G9, G11-G13 (no G10). "G2" = 1 item on the G side; "G2-L4" = 4 items on the L side only.
//    A 1-2 letter prefix is ignored: "MG3-L1" counts as G3 with a loss of 1.
//  - When a waybill appears in several files, the newest file wins, then the last row in it.
(function (root) {
  "use strict";

  const IN_EVENTS = ["جاهزة للالتقاط", "استلام من العميل"];
  const GRADES = ["G1", "G2", "G3", "G4", "G5", "G6", "G7", "G8", "G9", "G11", "G12", "G13"];
  const NO_STATUS = "(no status yet)";
  const NO_CONSIGNEE = "(no consignee yet)";
  const NO_PALLET = "(no pallet group)";

  const ALIASES = {
    Waybill: ["بوليصة الشحن", "رقم التتبع", "رقم البوليصة", "Waybill", "AWB", "Tracking Number", "Tracking No"],
    ReferenceNo: ["رقم المرجع", "الرقم المرجعي", "رقم المرجعي", "Reference", "Reference No", "Ref"],
    Contents: ["وصف محتويات الشحنة", "Contents", "Description", "Shipment Contents"],
    SecondaryRef: ["رقم المرجع الثانوي", "الرقم المرجعي الثانوي", "رقم المرجعي الثانوي", "Secondary Reference", "Secondary Ref", "Pallet", "Pallet Group"],
    CityName: ["مدينة الوجهة", "Destination City", "City"],
    LastEvent: ["آخر حدث", "Last Event", "Status"],
    ConsigneeName: ["اسم المرسل إليه", "Consignee", "Consignee Name", "Receiver"],
  };
  const LOGICAL = Object.keys(ALIASES);

  function normText(v) {
    if (v === null || v === undefined) return "";
    return String(v)
      .normalize("NFKC") // Arabic presentation forms (ﺟﺎﻫﺰﺓ, ﻻ) -> normal letters
      .replace(/[​-‏‪-‮⁦-⁩ـ﻿]/g, "") // invisible marks, tatweel
      .replace(/[\r\n\t ]/g, " ")
      .replace(/ {2,}/g, " ")
      .trim();
  }

  // Loose matching key: also drops Arabic diacritics and every space.
  function normKey(v) {
    return normText(v)
      .replace(/[ً-ٰٟ]/g, "")
      .replace(/ /g, "")
      .replace(/[آأإٱ]/g, "ا")
      .replace(/[ىی]/g, "ي")
      .replace(/ة/g, "ه")
      .toLowerCase();
  }

  // Fast path for the usual plain ASCII waybill (EBUXSL2082091026).
  function normWaybill(v) {
    if (v == null) return "";
    const s = String(v);
    if (/^[A-Za-z0-9-]+$/.test(s)) return s.toUpperCase();
    return normText(s).toUpperCase().replace(/ /g, "");
  }

  const HEADER_MAP = new Map();
  for (const logical of LOGICAL) for (const h of ALIASES[logical]) HEADER_MAP.set(normKey(h), logical);
  // IN events, as word fragments that must all appear (after normKey):
  //   جاهزة للالتقاط  (also typed للاتقاط / للإلتقاط)  and  استلام من العميل
  const IN_FRAGMENTS = [["جاهز", "تقاط"], ["استلام", "عميل"]];
  // Words that turn an event into its opposite ("لم يتم الاستلام من العميل").
  const NEGATIONS = ["لم", "عدم", "فشل", "تعذر", "الغاء", "ملغي", "ملغى", "مرفوض", "رفض"];

  const IN_FRAG_KEYS = IN_FRAGMENTS.map((frags) => frags.map(normKey));
  const NEG_KEYS = NEGATIONS.map((n) => ({ key: normKey(n), prefix: n.length > 2 }));
  const inCache = new Map();

  function isInEvent(event) {
    const raw = event == null ? "" : String(event);
    let v = inCache.get(raw);
    if (v !== undefined) return v;
    const k = normKey(raw);
    const words = normText(raw).split(" ").map(normKey);
    const negated = words.some((w) => NEG_KEYS.some((n) => w === n.key || (n.prefix && w.startsWith(n.key))));
    v = !!k && !negated && IN_FRAG_KEYS.some((frags) => frags.every((f) => k.includes(f)));
    inCache.set(raw, v);
    return v;
  }

  // Exact alias first; fuzzy rules only for columns not found exactly.
  function headerLogical(cell) {
    return HEADER_MAP.get(normKey(cell)) || null;
  }

  function fuzzyHeader(cell) {
    const k = normKey(cell);
    if (!k) return null;
    if (k.includes("مرجع")) return k.includes("ثانوي") ? "SecondaryRef" : "ReferenceNo";
    if (k.includes("بوليص") || k.includes("تتبع")) return "Waybill";
    if (k.includes("حدث") && !k.includes("تاريخ") && !k.includes("وقت")) return "LastEvent";
    if (k.includes("مرسلاليه")) return "ConsigneeName";
    if (k.includes("محتوي")) return "Contents";
    return null;
  }

  // Last grade code in the text: "… \\ G1", "… / G1", "… // G1-L4", "ASR-SEP30-G9".
  const gradeCache = new Map();
  function parseGradeLoss(text) {
    const raw = text == null ? "" : String(text);
    if (gradeCache.has(raw)) return gradeCache.get(raw);
    const v = parseGradeLossUncached(raw);
    if (gradeCache.size > 200000) gradeCache.clear();
    gradeCache.set(raw, v);
    return v;
  }

  function parseGradeLossUncached(text) {
    const s = normText(text);
    if (!s) return null;
    // optional 1-2 letter prefix (MG3), the grade, then optional -L<n>; never inside a longer word
    const re = /(?:^|[^A-Za-z0-9])[A-Za-z]{0,2}?G(\d{1,2})(?:\s*-\s*L\s*(\d{1,5}))?(?![A-Za-z0-9])/gi;
    let m, last = null;
    while ((m = re.exec(s)) !== null) last = m;
    if (!last) return null;
    const grade = "G" + parseInt(last[1], 10);
    if (!GRADES.includes(grade)) return null;
    return { grade, loss: last[2] ? parseInt(last[2], 10) : 0 };
  }

  // "<school> - مسارات - M174562" -> "M174562"
  const codeCache = new Map();
  function customerCode(name) {
    const raw = name == null ? "" : String(name);
    if (!codeCache.has(raw)) codeCache.set(raw, customerCodeUncached(raw));
    return codeCache.get(raw);
  }

  function customerCodeUncached(name) {
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
      if ((aoa[r] || []).some((c) => (headerLogical(c) || fuzzyHeader(c)) === "Waybill")) hr = r;
    }
    if (hr < 0) return null;
    const colOf = {};
    const header = aoa[hr] || [];
    header.forEach((c, i) => {
      const l = headerLogical(c);
      if (l && !(l in colOf)) colOf[l] = i;
    });
    header.forEach((c, i) => {
      const l = fuzzyHeader(c);
      if (l && !(l in colOf) && !Object.values(colOf).includes(i)) colOf[l] = i;
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
    return { kind, columns: Object.keys(colOf), headers: Object.fromEntries(Object.entries(colOf).map(([l, i]) => [l, normText(header[i])])), rows };
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
      return { kind: res.kind, sheet: name, headers: res.headers, rows: res.rows };
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
        // Rows are already cleaned by readSheet; keep a reference plus its file.
        if (f.kind === "master") {
          master.set(wb, { r, f });
        } else if (f.kind === "status" && r.LastEvent) {
          status.set(wb, { r, f });
        }
      }
    }

    const shipments = [];
    for (const [wb, mm] of master) {
      const m = mm.r;
      const gl = parseGradeLoss(m.Contents) || parseGradeLoss(m.SecondaryRef);
      const ss = status.get(wb);
      const s = ss && ss.r;
      const loss = gl ? gl.loss : 0;
      const inW = s ? isInEvent(s.LastEvent) : true;
      shipments.push({
        waybill: wb,
        reference: m.ReferenceNo || wb,
        secondaryRef: m.SecondaryRef,
        pallet: m.SecondaryRef || NO_PALLET,
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
        masterFile: mm.f.name,
        statusFile: s ? ss.f.name : "",
        statusDate: s ? ss.f.lastModified : null,
      });
    }
    const unmatched = [];
    for (const [wb, { r: s, f }] of status) {
      if (!master.has(wb)) {
        unmatched.push({ waybill: wb, lastEvent: s.LastEvent, city: s.CityName, consignee: s.ConsigneeName, file: f.name, fileDate: f.lastModified });
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
    let last = null;
    for (const s of shipments) if (s.statusDate && (last === null || s.statusDate > last)) last = s.statusDate;
    const pallets = arguments[2] || palletGroups(shipments);
    return {
      total, inCount, outCount: total - inCount,
      inPct: total ? inCount / total : 0,
      unmatched: unmatched.length,
      lastStatusDate: last,
      pallets: pallets.length,
      palletsWithRemaining: pallets.filter((p) => p.remaining > 0).length,
    };
  }

  // Pallet groups (الرقم المرجعي الثانوي): items in each group, how many are still IN.
  // items/in/out/remaining count G rows (1 each); loss is the L side.
  function palletGroups(shipments) {
    const m = new Map();
    for (const s of shipments) {
      let p = m.get(s.pallet);
      if (!p) {
        p = { pallet: s.pallet, items: 0, inCount: 0, outCount: 0, remaining: 0, loss: 0, lossOut: 0, waybills: 0,
              grades: new Set(), consignees: new Set(), lastUpdate: null };
        m.set(s.pallet, p);
      }
      p.waybills++;
      p.items += s.gQty;
      p.loss += s.lossQty;
      if (s.inWarehouse) p.inCount += s.gQty;
      else { p.outCount += s.gQty; p.lossOut += s.lossQty; }
      if (s.grade) p.grades.add(s.grade);
      if (s.hasStatus) p.consignees.add(s.consignee);
      if (s.statusDate && (!p.lastUpdate || s.statusDate > p.lastUpdate)) p.lastUpdate = s.statusDate;
    }
    const out = [];
    for (const p of m.values()) {
      p.remaining = p.items - p.outCount;
      p.status = p.items === 0 ? "Loss only" : p.outCount === 0 ? "Not started" : p.remaining === 0 ? "Complete" : "Partial";
      p.pctOut = p.items ? p.outCount / p.items : 0;
      p.grades = [...p.grades].sort((a, b) => parseInt(a.slice(1), 10) - parseInt(b.slice(1), 10));
      p.consignees = p.consignees.size;
      out.push(p);
    }
    return out.sort((a, b) => b.remaining - a.remaining || a.pallet.localeCompare(b.pallet));
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

  const api = { IN_EVENTS, GRADES, NO_STATUS, NO_CONSIGNEE, NO_PALLET, palletGroups, fuzzyHeader, normText, normKey, normWaybill, headerLogical, isInEvent,
    parseGradeLoss, customerCode, readSheet, parseWorkbook, compute, gradeStock, summary, byEvent, consigneeMatrix };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Inventory = api;
})(typeof window !== "undefined" ? window : globalThis);
