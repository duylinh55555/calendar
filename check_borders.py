import openpyxl

file_path = 'd:\\VS Code Projects\\calendar\\sample.xlsx'
wb = openpyxl.load_workbook(file_path)
ws = wb.active

for r in [1, 5, 7, 160, 163, 174]:
    print(f"\nRow {r} borders:")
    for c in range(1, 10):
        cell = ws.cell(row=r, column=c)
        if cell.value or cell.border:
            b = cell.border
            has_border = b and (b.top.style or b.bottom.style or b.left.style or b.right.style)
            print(f"Col {c} ('{cell.value}'): has_border={has_border}")

