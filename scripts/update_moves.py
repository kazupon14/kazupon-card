import os, csv, openpyxl, qrcode

XLSX = "カード図鑑.xlsx"
SHEET = "技図鑑"
CSV_OUT = "moves.csv"
QR_DIR = "技QRコード"
BASE_URL = "https://kazupon14.github.io/kazupon-card/move.html?id="

wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb[SHEET]

headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
if "技ID" not in headers:
    raise SystemExit("技ID列が見つかりません。")

id_col = headers.index("技ID") + 1
rows = []
ids = []

for r in range(2, ws.max_row + 1):
    vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
    tech_id = ws.cell(r, id_col).value
    if tech_id:
        tid = str(tech_id).strip()
        ids.append(tid)
        rows.append(vals)

with open(CSV_OUT, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f)
    w.writerow(headers)
    w.writerows(rows)

os.makedirs(QR_DIR, exist_ok=True)
created = []
for tech_id in ids:
    path = os.path.join(QR_DIR, f"{tech_id}.png")
    if os.path.exists(path):
        continue

    url = BASE_URL + tech_id
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=12,
        border=4,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img.save(path)
    created.append(tech_id)

print(f"moves.csv 更新完了: {len(ids)}技")
print(f"新規QR: {len(created)}件")
if created:
    print("追加:", ", ".join(created))
