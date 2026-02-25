import xlrd
import os

project_root = os.path.dirname(os.path.dirname(__file__))
file_path = os.path.join(project_root, 'sample.xls')

print(f"Opening {file_path}")
workbook = xlrd.open_workbook(file_path, formatting_info=True)
sheet = workbook.sheet_by_index(0)

print("Merged cells:", sheet.merged_cells[:5])

# check a cell
for r in range(min(sheet.nrows, 15)):
    for c in range(min(sheet.ncols, 15)):
        val = sheet.cell(r, c).value
        if val and str(val).strip():
            xf_idx = sheet.cell_xf_index(r, c)
            xf = workbook.xf_list[xf_idx]
            font = workbook.font_list[xf.font_index]
            is_bold = font.weight >= 700
            print(f"Row {r} Col {c}: {val}, bold={is_bold}, color_idx={font.colour_index}")
            
# get rgb map
rgb = workbook.colour_map.get(10) # just an example
print(f"Color map example: {rgb}")
