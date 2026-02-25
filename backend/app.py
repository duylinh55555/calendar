import openpyxl
from flask import Flask, jsonify, request
from flask_cors import CORS
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
# Ensure JSON responses are sent with UTF-8 encoding for Unicode characters
app.json.ensure_ascii = False
CORS(app)  # This will enable CORS for all routes

# Config upload folder
UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/api/upload', methods=['POST'])
def upload_schedule():
    if 'file' not in request.files:
        return jsonify({"error": "Không có file nào được chọn."}), 400
    file = request.files['file']
    week = request.form.get('week')

    if file.filename == '':
        return jsonify({"error": "Tên file trống."}), 400
    if not week:
        return jsonify({"error": "Vui lòng chọn tuần."}), 400

    if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
        # Use week string as filename safely
        safe_week = week.replace('/', '_').replace('\\', '_')
        filename = f"{safe_week}.xlsx"
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        try:
            file.save(file_path)
            return jsonify({"message": f"Tải lên thành công lịch cho {week}"}), 200
        except Exception as e:
            return jsonify({"error": f"Lỗi khi lưu file: {str(e)}"}), 500
    else:
        return jsonify({"error": "Chỉ chấp nhận file định dạng .xlsx hoặc .xls"}), 400

@app.route('/api/schedule')
def get_schedule():
    week = request.args.get('week')
    try:
        # Construct the path to the Excel file relative to the script's location
        script_dir = os.path.dirname(__file__)  # a.k.a. 'backend' folder
        project_root = os.path.dirname(script_dir) # a.k.a. 'calendar' folder
        
        if week:
            safe_week = week.replace('/', '_').replace('\\', '_')
            filename = f"{safe_week}.xlsx"
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            if not os.path.exists(file_path):
                return jsonify({"error": f"Không tìm thấy dữ liệu lịch cho '{week}'. Vui lòng tải lên file lịch cho tuần này."}), 404
        else:
            file_path = os.path.join(project_root, 'sample.xlsx')

        # Use openpyxl to get merged cells info
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active

        # Build a dictionary of how each cell is merged
        merged_cells_info = {}
        for merged_range in sheet.merged_cells.ranges:
            min_col, min_row, max_col, max_row_merge = merged_range.bounds
            for r in range(min_row, max_row_merge + 1):
                for c in range(min_col, max_col + 1):
                    if r == min_row and c == min_col:
                        merged_cells_info[(r, c)] = {
                            'row_span': max_row_merge - min_row + 1,
                            'col_span': max_col - min_col + 1
                        }
                    else:
                        merged_cells_info[(r, c)] = {
                            'row_span': 0,
                            'col_span': 0
                        }

        # Detect start row (look for "TT" or similar header)
        start_row = 1
        for r in range(1, min(sheet.max_row, 50) + 1):
            found_header = False
            for c in range(1, min(sheet.max_column, 20) + 1):
                val = sheet.cell(row=r, column=c).value
                if isinstance(val, str) and val.strip().upper() == 'TT':
                    start_row = r
                    found_header = True
                    break
            if found_header:
                break
                
        # Detect actual max column in the table
        max_col = 1
        for c in range(1, sheet.max_column + 1):
            val = sheet.cell(row=start_row, column=c).value
            if val is not None and str(val).strip() != "":
                max_col = max(max_col, c)
        
        # In case the table starts with merged cells that span further
        for (r, c), info in merged_cells_info.items():
            if r == start_row and info['row_span'] > 0:
                max_col = max(max_col, c + info['col_span'] - 1)
                
        # To find max_col reliably based on data across a few header rows
        for r in range(start_row, min(start_row + 3, sheet.max_row + 1)):
            for c in range(1, sheet.max_column + 1):
                if sheet.cell(row=r, column=c).value is not None:
                    max_col = max(max_col, c)
                    
        # Detect end row (ignore signatures)
        # We find the last row that either has a border or has substantial data in the table column width
        end_row = start_row
        consecutive_empty = 0
        for r in range(start_row, sheet.max_row + 1):
            row_is_part_of_table = False
            
            for c in range(1, max_col + 1):
                cell = sheet.cell(row=r, column=c)
                # Check if it has a border (strong indicator of table structure)
                has_border = cell.border and (cell.border.top.style or cell.border.bottom.style or cell.border.left.style or cell.border.right.style)
                if has_border:
                    row_is_part_of_table = True
                    break
                    
                # Check if it's a merged cell connected to previous data
                merge_info = merged_cells_info.get((r, c))
                if merge_info and (merge_info['row_span'] > 0 or merge_info['row_span'] == 0): # It's part of a merge
                    # Specifically, if it's a secondary cell in a merge, it means the top cell was part of the table
                    # Actually, if any cell is part of a merge that started at or before this row and spans into it.
                    pass
            
            # If no border, check if there's text but only in few columns (like signatures at the end)
            # Signatures usually don't have borders in these templates.
            if row_is_part_of_table:
                end_row = r
                consecutive_empty = 0
            else:
                # If it has data, check if it's a signature string
                has_data = False
                for c in range(1, max_col + 1):
                    val = sheet.cell(row=r, column=c).value
                    if val is not None and str(val).strip() != "":
                        has_data = True
                        break
                
                if has_data:
                    # Let's see if it's just a signature
                    val_str = " ".join([str(sheet.cell(row=r, column=c_idx).value) for c_idx in range(1, max_col + 1) if sheet.cell(row=r, column=c_idx).value is not None])
                    val_str_up = val_str.upper()
                    if "TRƯỞNG BAN" in val_str_up or "CHỮ KÝ" in val_str_up or "ĐẠI TÁ" in val_str_up or "HIỆU TRƯỞNG" in val_str_up or "NGÀY" in val_str_up:
                        break # Stop at signature
                    else:
                        end_row = r # Might be table data without borders
                        consecutive_empty = 0
                else:
                    consecutive_empty += 1
                    if consecutive_empty > 3:
                        break # 3 empty rows probably mean end of table
        
        # Also expand max_col if a merged cell containing data extends further within table bounds
        for (r, c), info in merged_cells_info.items():
            if start_row <= r <= end_row and info['row_span'] > 0: 
                cell = sheet.cell(row=r, column=c)
                if cell.value is not None:
                    max_col = max(max_col, c + info['col_span'] - 1)

        # Store cell data (value, rowspan, colspan)
        data = []
        for r in range(start_row, end_row + 1):
            row_data = []
            for c in range(1, max_col + 1):
                cell = sheet.cell(row=r, column=c)
                
                merge_info = merged_cells_info.get((r, c))
                if merge_info:
                    row_span = merge_info['row_span']
                    col_span = merge_info['col_span']
                else:
                    row_span = 1
                    col_span = 1
                
                # Truncate rowspan and colspan if they exceed end_row and max_col
                if row_span > 0:
                    if r + row_span - 1 > end_row:
                        row_span = end_row - r + 1
                    if c + col_span - 1 > max_col:
                        col_span = max_col - c + 1
                
                # Append cell data if it's not a secondary cell in a merge
                if row_span != 0:
                    # ensure cell value is at least string if not none
                    cell_val = "" if cell.value is None else str(cell.value).strip()
                    
                    font_color = None
                    is_bold = False
                    try:
                        if cell.font:
                            is_bold = bool(cell.font.b)
                            if cell.font.color and hasattr(cell.font.color, 'rgb') and cell.font.color.rgb:
                                rgb_val = str(cell.font.color.rgb)
                                if rgb_val.startswith('FF') and len(rgb_val) == 8:
                                    font_color = '#' + rgb_val[2:]
                                elif len(rgb_val) == 6 and rgb_val != '000000':
                                    font_color = '#' + rgb_val
                    except Exception:
                        pass

                    row_data.append({
                        "value": cell_val,
                        "rowspan": row_span,
                        "colspan": col_span,
                        "is_bold": is_bold,
                        "font_color": font_color
                    })

            data.append(row_data)

        return jsonify(data)
    except FileNotFoundError:
        return jsonify({"error": "The file 'sample.xls' was not found."}), 404
    except Exception as e:
        # It's good practice to log the actual exception
        app.logger.error(f"An error occurred: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    # Running on port 5001 to avoid potential conflicts with other services
    app.run(debug=True, port=5001)
