import os
import re
import json
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

columns = [
    'no', 'memo', 'idp', 'dept', 'deptname', 'status', 'name', 'addr', 'tel', 'tmb',
    'p1', 'p2', 'sum_p', 'b1', 'b2', 'b3', 'b4', 'b5', 'b6', 'b71', 'b72',
    'd1', 'd2', 'd3', 'd4', 'd5', 'd6', 'd7', 'd81', 'd82', 'd83', 'd84', 't1', 't2'
]

col_th_titles = {
    'no': 'ลำดับ',
    'memo': 'เลขสมาชิก',
    'idp': 'เลขบัตรประชาชน',
    'dept': 'รหัสหน่วย',
    'deptname': 'ชื่อหน่วย',
    'status': 'สถานะ',
    'name': 'ยศ-ชื่อ-สกุล',
    'addr': 'ที่อยู่',
    'tel': 'เบอร์โทร',
    'tmb': 'เลขบัญชี TTB',
    'p1': 'เงินปันผล (5.02%)',
    'p2': 'เงินเฉลี่ยคืน (13.50%)',
    'sum_p': 'รวมรับเงินปันผลเฉลี่ยคืน',
    'b1': 'ทุนเรือนหุ้นสะสม',
    'b2': 'เงินกู้สามัญคงเหลือ',
    'b3': 'เงินกู้พิเศษ1คงเหลือ',
    'b4': 'เงินกู้พิเศษ2คงเหลือ',
    'b5': 'เงินกู้พิเศษ3คงเหลือ',
    'b6': 'เงินกู้พิเศษ4คงเหลือ',
    'b71': 'ภาระค้ำประกันคงเหลือ 1',
    'b72': 'ภาระค้ำประกันคงเหลือ 2',
    'd1': 'หัก เงินประกันชีวิต กรมธรรม์ที่ 3',
    'd2': 'หัก เงินประกันชีวิต เงินกู้สามัญ',
    'd3': 'หัก เงินประกันภัย ผู้ค้ำประกัน',
    'd4': 'หัก เงินประกันชีวิต เงินกู้พิเศษ',
    'd5': 'หัก ชำระหนี้ ภาระค้ำประกัน',
    'd6': 'หัก ปันผล (ชำระต้นเงินสามัญ)',
    'd7': 'หัก ชำระหนี้อื่นๆ',
    'd81': 'หัก ฌาปนกิจ สสอท.',
    'd82': 'หัก ฌาปนกิจ สส.ชสอ.',
    'd83': 'หัก ฌาปนกิจ สส.สท.',
    'd84': 'หัก ฌาปนกิจ สส.สก.',
    't1': 'รวมหักเงินทั้งสิ้น',
    't2': 'ยอดรับสุทธิ'
}

def parse_sql_values(row_str):
    """Parse a single SQL tuple string into list of values."""
    row_str = row_str.strip()
    if row_str.startswith('('):
        row_str = row_str[1:]
    if row_str.endswith('),') or row_str.endswith(');'):
        row_str = row_str[:-2]
    elif row_str.endswith(')'):
        row_str = row_str[:-1]

    values = []
    i = 0
    n = len(row_str)
    while i < n:
        # Skip whitespace
        while i < n and row_str[i] in ' \t':
            i += 1
        if i >= n:
            break

        if row_str[i] == "'":
            # Quoted string
            i += 1
            s = []
            while i < n:
                if row_str[i] == '\\' and i + 1 < n:
                    s.append(row_str[i+1])
                    i += 2
                elif row_str[i] == "'":
                    if i + 1 < n and row_str[i+1] == "'":
                        s.append("'")
                        i += 2
                    else:
                        i += 1
                        break
                else:
                    s.append(row_str[i])
                    i += 1
            values.append("".join(s))
            # skip to next comma
            while i < n and row_str[i] != ',':
                i += 1
            if i < n and row_str[i] == ',':
                i += 1
        else:
            # Unquoted value (number or NULL)
            start = i
            while i < n and row_str[i] != ',':
                i += 1
            val = row_str[start:i].strip()
            if i < n and row_str[i] == ',':
                i += 1
            if val.upper() == 'NULL':
                values.append(None)
            else:
                try:
                    if '.' in val:
                        values.append(float(val))
                    else:
                        values.append(int(val))
                except ValueError:
                    values.append(val)

    return values

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    sql_path = os.path.join(base_dir, '..', 'rtarfcoo_infodiv.sql')
    output_excel = os.path.join(base_dir, 'data2025.xlsx')
    output_json = os.path.join(base_dir, 'data2025.json')

    print(f"Reading SQL file from: {sql_path}")
    with open(sql_path, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()

    all_data = []
    for line_num, line in enumerate(lines, 1):
        line = line.strip()
        if line.startswith('(') and (line.endswith('),') or line.endswith(');')):
            vals = parse_sql_values(line)
            if len(vals) == len(columns):
                row_dict = dict(zip(columns, vals))
                all_data.append(row_dict)
            else:
                print(f"Warning: Line {line_num} had {len(vals)} values instead of {len(columns)}")

    print(f"Parsed {len(all_data)} records successfully.")

    # Save to JSON for web application usage
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(all_data, f, ensure_ascii=False, indent=2)
    print(f"Saved JSON to: {output_json}")

    # Build DataFrame
    df = pd.DataFrame(all_data)

    # Ensure memo and idp and tmb are stored as formatted strings
    df['memo'] = df['memo'].apply(lambda x: f"{int(x):05d}" if pd.notnull(x) and str(x).isdigit() else str(x or ''))
    df['idp'] = df['idp'].apply(lambda x: str(x or ''))
    df['tmb'] = df['tmb'].apply(lambda x: str(x or ''))
    df['tel'] = df['tel'].apply(lambda x: str(x or ''))

    # Rename columns for Excel
    df_excel = df.rename(columns=col_th_titles)

    # Write to Excel using openpyxl with professional styling
    print("Writing to Excel with openpyxl...")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ข้อมูลสมาชิกปี 2568"

    # Header style
    header_font = Font(name='TH Sarabun New', size=14, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='1E3A8A', end_color='1E3A8A', fill_type='solid') # Navy Blue
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    data_font = Font(name='TH Sarabun New', size=13)
    thin_border = Border(
        left=Side(style='thin', color='D1D5DB'),
        right=Side(style='thin', color='D1D5DB'),
        top=Side(style='thin', color='D1D5DB'),
        bottom=Side(style='thin', color='D1D5DB')
    )

    # Write headers
    headers = list(df_excel.columns)
    ws.append(headers)
    ws.row_dimensions[1].height = 28

    for col_idx, cell in enumerate(ws[1], 1):
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    # Write data rows
    text_cols = {'memo', 'idp', 'tel', 'tmb', 'deptname', 'status', 'name', 'addr'}
    decimal_cols = {'b2', 'b71', 'b72', 'd3', 'd5', 'd7', 't1', 't2'}
    int_cols = {'no', 'dept', 'p1', 'p2', 'sum_p', 'b1', 'b3', 'b4', 'b5', 'b6', 'd1', 'd2', 'd4', 'd6', 'd81', 'd82', 'd83', 'd84'}

    for row_idx, row in df.iterrows():
        excel_row = []
        for col_name in columns:
            val = row[col_name]
            excel_row.append(val)
        ws.append(excel_row)

        curr_row = row_idx + 2
        ws.row_dimensions[curr_row].height = 20

        # Alternate row background
        row_fill = PatternFill(start_color='F9FAFB', end_color='F9FAFB', fill_type='solid') if row_idx % 2 == 1 else None

        for col_idx, col_name in enumerate(columns, 1):
            cell = ws.cell(row=curr_row, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            if row_fill:
                cell.fill = row_fill

            if col_name in text_cols:
                cell.number_format = '@'
                cell.alignment = Alignment(horizontal='left', vertical='center')
                if col_name in ['memo', 'idp', 'tel', 'tmb']:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
            elif col_name in decimal_cols:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif col_name in int_cols:
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal='right', vertical='center')
                if col_name in ['no', 'dept']:
                    cell.alignment = Alignment(horizontal='center', vertical='center')

    # Freeze header row
    ws.freeze_panes = 'C2'

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        header_val = str(col[0].value or '')
        # Approximate width based on header length
        max_len = max(len(header_val) * 1.5, 10)
        ws.column_dimensions[col_letter].width = min(max_len + 3, 50)

    wb.save(output_excel)
    print(f"Successfully exported {len(all_data)} rows to {output_excel}")

if __name__ == '__main__':
    main()
