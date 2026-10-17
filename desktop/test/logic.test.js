// node test/logic.test.js  - checks the shared logic against the sample files.
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const L = require("../app/logic.js");

// grade parser
const cases = [
  ["نقطة توزيع المجاردة \\\\ G9-L1", "G9", 1], ["x / G1", "G1", 0], ["x // G1-L4", "G1", 4],
  ["ASR-SEP30-G9", "G9", 0], ["G13 - L 2", "G13", 2], ["x \\\\ G10", null],
  ["g2-l300", "G2", 300], ["G2-L", "G2", 0], ["G1-L4x", "G1", 0],
];
for (const [t, g, l] of cases) {
  const r = L.parseGradeLoss(t);
  if (g === null) assert.strictEqual(r, null, t);
  else assert.deepStrictEqual(r, { grade: g, loss: l }, t);
}
assert.strictEqual(L.customerCode("ثانوية منار العلم الأهلية - مسارات - M174562"), "M174562");
assert.strictEqual(L.headerLogical(" اخر  حدث "), "LastEvent"); // alef variant + spaces

// real .xlsx samples through SheetJS
const dir = path.join(__dirname, "..", "..", "access", "samples");
const files = ["master_inventory_sample.xlsx", "status_update_sample.xlsx"].map((n, i) => {
  const res = L.parseWorkbook(XLSX, fs.readFileSync(path.join(dir, n)));
  assert(!res.error, res.error);
  return { name: n, lastModified: i, ...res };
});
assert.strictEqual(files[0].kind, "master");
assert.strictEqual(files[1].kind, "status");

const { shipments, unmatched } = L.compute(files);
const sum = L.summary(shipments, unmatched);
assert.deepStrictEqual([sum.total, sum.inCount, sum.outCount, sum.unmatched], [9, 4, 5, 1]);
assert.strictEqual(unmatched[0].waybill, "EBUXSL2079999999");

const gs = L.gradeStock(shipments);
const g1 = gs.rows.find((r) => r.grade === "G1");
const g9 = gs.rows.find((r) => r.grade === "G9");
assert.deepStrictEqual([g1.gTotal, g1.gIn, g1.gOut, g1.gRem, g1.lTotal, g1.lIn, g1.lOut, g1.lRem], [5, 2, 3, 2, 4, 0, 4, 0]);
assert.deepStrictEqual([g9.gTotal, g9.gIn, g9.gOut, g9.gRem, g9.lTotal, g9.lIn, g9.lOut, g9.lRem], [2, 2, 0, 2, 1, 0, 1, 0]);
assert.strictEqual(gs.rows.length, 12);
assert(!gs.rows.some((r) => r.grade === "G10"));

const mx = L.consigneeMatrix(shipments);
assert.strictEqual(mx.consignees.length, 3);

// newest file wins
const older = { name: "a", lastModified: 1, kind: "status", rows: [{ Waybill: "EBUXSL2079600003", LastEvent: "جاهزة للالتقاط" }] };
const newer = { name: "b", lastModified: 2, kind: "status", rows: [{ Waybill: "EBUXSL2079600003", LastEvent: "التسليم النهائي" }] };
const r2 = L.compute([files[0], newer, older]).shipments.find((s) => s.waybill === "EBUXSL2079600003");
assert.strictEqual(r2.location, "OUT");

console.log("All logic tests passed.");

// --- real-file cases reported from the app ---
assert.deepStrictEqual(L.parseGradeLoss("نقطة توزيع شرورة \\\\ MG3-L1"), { grade: "G3", loss: 1 });
assert.deepStrictEqual(L.parseGradeLoss("نقطة توزيع \\\\ MG3"), { grade: "G3", loss: 0 });
assert.deepStrictEqual(L.parseGradeLoss("G05"), { grade: "G5", loss: 0 });
assert.strictEqual(L.parseGradeLoss("x \\\\ MG10-L2"), null);
assert.strictEqual(L.parseGradeLoss("BIGG5"), null);
for (const t of ["جاهزة للالتقاط", "ﺟﺎﻫﺰﺓ ﻟﻼﻟﺘﻘﺎﻁ", "جاهزةٌ للالتقاط", "جاهزة‌ للالتقاط", "جاهزة للإلتقاط",
                 "جاهزه  للالتقاط ", "استلام من العميل", "جاهزة للالتقاط - فرع نجران"]) assert(L.isInEvent(t), t);
for (const t of ["التسليم النهائي", "قيد التوصيل", ""]) assert(!L.isInEvent(t), t);
console.log("Real-file cases passed.");

// --- real headers: الرقم المرجعي / الرقم المرجعي الثانوي ---
{
  const r = L.readSheet([
    ["بوليصة الشحن", "الرقم المرجعي", "وصف محتويات الشحنة", "الرقم المرجعي الثانوي"],
    ["EBUXSL2082091026", "26524-2", "نقطة توزيع شرورة \\\\ MG3-L1", "MKK-SEP30CH"],
  ]);
  assert.strictEqual(r.kind, "master");
  assert.strictEqual(r.rows[0].ReferenceNo, "26524-2");
  assert.strictEqual(r.rows[0].SecondaryRef, "MKK-SEP30CH");
  // unknown wording still found by the fuzzy rules
  const f = L.readSheet([["رقم البوليصه", "المرجع", "مرجع ثانوي للشحنة", "المحتويات"], ["X1", "1-1", "P1", "\\\\ G2"]]);
  assert.deepStrictEqual([f.rows[0].Waybill, f.rows[0].ReferenceNo, f.rows[0].SecondaryRef, f.rows[0].Contents], ["X1", "1-1", "P1", "\\\\ G2"]);
}
// IN spelling variants and negations
for (const t of ["جاهزة للاتقاط", "جاهز للالتقاط", "تم الاستلام من العميل"]) assert(L.isInEvent(t), t);
for (const t of ["لم يتم الاستلام من العميل", "فشل الاستلام من العميل", "تعذر الاستلام من العميل"]) assert(!L.isInEvent(t), t);

// --- pallet groups ---
{
  const files = [
    { name: "m", lastModified: 1, kind: "master", rows: [
      { Waybill: "A1", ReferenceNo: "1", Contents: "x \\\\ G3", SecondaryRef: "MKK-SEP30CH" },
      { Waybill: "A2", ReferenceNo: "2", Contents: "x \\\\ G3", SecondaryRef: "MKK-SEP30CH" },
      { Waybill: "A3", ReferenceNo: "3", Contents: "x \\\\ MG3-L1", SecondaryRef: "MKK-SEP30CH" },
      { Waybill: "B1", ReferenceNo: "4", Contents: "x \\\\ G5", SecondaryRef: "PAL-2" },
      { Waybill: "C1", ReferenceNo: "5", Contents: "x \\\\ G5", SecondaryRef: "" },
    ] },
    { name: "s", lastModified: 2, kind: "status", rows: [
      { Waybill: "A1", LastEvent: "التسليم النهائي" },
      { Waybill: "A2", LastEvent: "جاهزة للاتقاط" },
      { Waybill: "A3", LastEvent: "التسليم النهائي" },
      { Waybill: "B1", LastEvent: "التسليم النهائي" },
    ] },
  ];
  const { shipments } = L.compute(files);
  const pg = Object.fromEntries(L.palletGroups(shipments).map((p) => [p.pallet, p]));
  const mkk = pg["MKK-SEP30CH"];
  assert.deepStrictEqual([mkk.items, mkk.inCount, mkk.outCount, mkk.remaining, mkk.loss, mkk.status], [2, 1, 1, 1, 1, "Partial"]);
  assert.strictEqual(pg["PAL-2"].status, "Complete");
  assert.strictEqual(pg[L.NO_PALLET].status, "Not started");
}

// --- scale: 300,000 waybills ---
{
  const N = 300000, mrows = [], srows = [];
  for (let i = 0; i < N; i++) {
    const g = [1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 13][i % 12];
    mrows.push({ Waybill: "EBUX" + i, ReferenceNo: "R" + i, Contents: `نقطة توزيع \\\\ G${g}${i % 50 === 0 ? "-L2" : ""}`, SecondaryRef: "PAL-" + (i % 5000) });
    if (i % 3) srows.push({ Waybill: "EBUX" + i, LastEvent: i % 3 === 1 ? "التسليم النهائي" : "جاهزة للالتقاط", CityName: "حائل", ConsigneeName: "مدرسة " + (i % 2000) + " - M" + (i % 2000) });
  }
  const t0 = Date.now();
  const { shipments, unmatched } = L.compute([{ name: "m", lastModified: 1, kind: "master", rows: mrows }, { name: "s", lastModified: 2, kind: "status", rows: srows }]);
  const sm = L.summary(shipments, unmatched);
  const gs = L.gradeStock(shipments);
  const pg = L.palletGroups(shipments);
  L.consigneeMatrix(shipments);
  const ms = Date.now() - t0;
  assert.strictEqual(sm.total, N);
  assert.strictEqual(sm.outCount, srows.filter((r) => r.LastEvent === "التسليم النهائي").length);
  assert.strictEqual(pg.length, 5000);
  assert.strictEqual(gs.total.gTotal + gs.total.lTotal / 2, N);
  console.log(`Scale test: ${N.toLocaleString()} waybills computed in ${ms} ms.`);
}
console.log("Pallet and header cases passed.");
