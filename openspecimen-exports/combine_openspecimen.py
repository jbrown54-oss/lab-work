import csv
import re
import sys

# openpyxl lets Python create Excel files. It's optional: without it, the
# script saves combined.csv instead, which you can open in Excel or Google Sheets.
try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter
except ImportError:
    Workbook = None

csv.field_size_limit(sys.maxsize)


def load(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    return rows[0], rows[1:]


def by_ppid(header, rows):
    i = header.index("PPID")
    groups = {}
    for row in rows:
        groups.setdefault(row[i].strip().upper(), []).append(row)
    return groups


p_head, p_rows = load("participants.csv")
v_head, v_rows = load("visits.csv")
c_head, c_rows = load("NSTB_consent_forms.csv")

visits = by_ppid(v_head, v_rows)
consents = by_ppid(c_head, c_rows)
p_key = p_head.index("PPID")

rows = []
for p in p_rows:
    k = p[p_key].strip().upper()
    for v in visits.get(k, [[""] * len(v_head)]):
        for c in consents.get(k, [[""] * len(c_head)]):
            rows.append(p + v + c)

header = p_head + v_head + c_head
illegal = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
rows = [[illegal.sub("", v) for v in row] for row in rows]

if Workbook is None:
    with open("combined.csv", "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(rows)
    print(f"combined.csv: {len(rows)} rows, {len(header)} columns (openpyxl not installed, saved as CSV)")
    sys.exit()

wb = Workbook()
ws = wb.active
ws.title = "combined"
ws.append(header)
for row in rows:
    ws.append([v or None for v in row])

for i, row in enumerate(ws.iter_rows(), start=1):
    for cell in row:
        cell.alignment = Alignment(vertical="top")
        if i == 1:
            cell.font = Font(bold=True)
            cell.number_format = "@"
        elif isinstance(cell.value, str):
            cell.number_format = "@"
            cell.data_type = "s"

ws.freeze_panes = "A2"
ws.auto_filter.ref = ws.dimensions
for i, name in enumerate(header, start=1):
    width = max([len(name)] + [len(r[i - 1]) for r in rows]) + 2
    ws.column_dimensions[get_column_letter(i)].width = min(max(width, 10), 45)

wb.save("combined.xlsx")
print(f"combined.xlsx: {len(rows)} rows, {len(header)} columns")
