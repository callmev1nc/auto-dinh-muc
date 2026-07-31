import openpyxl, sys, io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

path = r"E:\Work\Thai-work\auto-dinh-muc\templates\Định mức - Bao giấy (KP - in offset).xlsx"
wb = openpyxl.load_workbook(path, data_only=True)

ws = wb["In"]

# Dump all non-None cells in "In" to understand its layout
print("=== In: all non-empty cells ===")
for r in range(1, 49):
    for c in range(1, 20):
        v = ws.cell(r, c).value
        if v is not None:
            col = chr(64 + c) if c <= 26 else f"col{c}"
            print(f"  {col}{r} = {v!r}")
