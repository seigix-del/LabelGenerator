"""Writes two small Excel files that mirror the real layouts, for testing.

    python access/tools/make_samples.py
"""
import pathlib

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

OUT = pathlib.Path(__file__).resolve().parent.parent / "samples"

MASTER_HEADERS = ["بوليصة الشحن", "رقم المرجع", "وصف محتويات الشحنة", "رقم المرجع الثانوي"]
STATUS_HEADERS = ["رقم التتبع", "مدينة الوجهة", "آخر حدث", "اسم المرسل إليه"]

READY = "جاهزة للالتقاط"           # IN
RECEIVED = "استلام من العميل"      # IN
DELIVERED = "التسليم النهائي"      # OUT
IN_TRANSIT = "قيد التوصيل"         # unknown -> auto-added as OUT

MASTER_ROWS = [
    ["EBUXSL2079602026", "71204-1", "نقطة توزيع المجاردة \\\\ G9-L1", "ASR-SEP30-G9"],
    ["EBUXSL2079418826", "71204-2", "نقطة توزيع المجاردة \\\\ G9-L1", "ASR-SEP30-G9"],
    ["EBUXSL2079500001", "71205-1", "نقطة توزيع حائل \\\\ G10-L2", "ASR-SEP30-G10"],
    ["EBUXSL2079500002", "71205-2", "نقطة توزيع حائل \\\\ G10-L2", "ASR-SEP30-G10"],
    ["EBUXSL2079500003", "71206-1", "نقطة توزيع أبها \\\\ G11-L1", "ASR-SEP30-G11"],
    ["EBUXSL2079500004", "71206-2", "نقطة توزيع أبها \\\\ G11-L1", "ASR-SEP30-G11"],
]

STATUS_ROWS = [
    ["EBUXSL2079418826", "حائل", DELIVERED, "ثانوية منار العلم الأهلية - مسارات - M174562"],
    ["EBUXSL2079602026", "المجاردة", READY, "ثانوية المجاردة - مسارات - M174001"],
    ["EBUXSL2079500001", "حائل", RECEIVED, "متوسطة حائل - مسارات - M174100"],
    ["EBUXSL2079500002", "حائل", IN_TRANSIT, "متوسطة حائل - مسارات - M174100"],
    ["EBUXSL2079500003", "أبها", DELIVERED, "ثانوية أبها - مسارات - M174200"],
    ["EBUXSL2079999999", "الرياض", DELIVERED, "مدرسة غير مسجلة - M999999"],   # not in master
    ["EBUXSL2079500003", "أبها", DELIVERED, "ثانوية أبها - مسارات - M174200"],  # duplicate row
]


def write(path, headers, rows):
    wb = Workbook()
    ws = wb.active
    ws.sheet_view.rightToLeft = True
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9D9D9")
    for row in rows:
        ws.append(row)
    for col in "ABCD":
        ws.column_dimensions[col].width = 32
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


if __name__ == "__main__":
    OUT.mkdir(exist_ok=True)
    write(OUT / "master_inventory_sample.xlsx", MASTER_HEADERS, MASTER_ROWS)
    write(OUT / "status_update_sample.xlsx", STATUS_HEADERS, STATUS_ROWS)
    print("written to", OUT)
