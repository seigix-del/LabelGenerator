// node test/logic.test.js  - checks the shared logic against the sample files.
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const L = require("../app/logic.js");

// grade parser
const cases = [
  ["نقطة توزيع المجاردة \\\\ G9-L1", "G9", 1], ["x / G1", "G1", 0], ["x // G1-L4", "G1", 4],
  ["ASR-SEP30-G9", "G9", 0], ["G13 - L 2", "G13", 2], ["PG13", null], ["x \\\\ G10", null],
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
