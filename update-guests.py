"""Rebuild guests.js from the seating Excel in this folder (sheet "מארחת לפי שמות משפחה")."""
import glob, json, re, sys
import openpyxl

SHEET = "מארחת לפי שמות משפחה"

files = glob.glob("*.xlsx")
if len(files) != 1:
    sys.exit(f"Expected exactly one .xlsx in this folder, found {len(files)}")

rows = openpyxl.load_workbook(files[0], data_only=True)[SHEET].iter_rows(values_only=True)
guests = []
for r in rows:
    last, first, table, count = r[1:5]
    if isinstance(last, str) and last.strip() and last.strip() != "שם משפחה" and isinstance(table, (int, float)):
        first = re.sub(r"\s*,\s*", ", ", str(first or "").strip())
        guests.append((last.strip(), first, int(table), int(count or 0)))

src = open("guests.js", encoding="utf-8").read()
head = src[: src.index("window.GUESTS = [")]
body = ",\n".join(
    f"  {{ last: {json.dumps(l, ensure_ascii=False)}, first: {json.dumps(f, ensure_ascii=False)}, table: {t} }}"
    for l, f, t, _ in guests
)
open("guests.js", "w", encoding="utf-8").write(head + "window.GUESTS = [\n" + body + ",\n];\n")
print(f"{len(guests)} invitations, {sum(g[3] for g in guests)} people, {len({g[2] for g in guests})} tables")
