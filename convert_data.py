"""
=============================================================================
ระบบแปลงข้อมูลเงินปันผลเฉลี่ยคืน สหกรณ์ออมทรัพย์ บก.ทท.
ไฟล์: convert_data.py
หน้าที่: อ่านข้อมูลจาก data-mast.xlsx แยกตามแต่ละ Sheet 
       โดยนำชื่อ Sheet มากำหนดเป็นปีข้อมูล และส่งออกเป็น data<year>.json และ data<year>.xlsx
=============================================================================
"""

import os
import sys
import re
import json
import shutil
import openpyxl

# Ensure UTF-8 output on Windows console
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

# =============================================================================
# 📋 แม่แบบโครงสร้างคอลัมน์ (COLUMN TEMPLATE MAPPING)
# กำหนดลำดับคอลัมน์ (Index 0-33) และชื่อฟิลด์ในระบบอย่างชัดเจน
# =============================================================================
COLUMN_TEMPLATE = {
    # --- ข้อมูลทั่วไปสมาชิก ---
    0:  {'field': 'no',       'title': 'ลำดับ',                            'type': 'int'},
    1:  {'field': 'memo',     'title': 'เลขสมาชิก',                         'type': 'memo'},
    2:  {'field': 'idp',      'title': 'เลขบัตรประชาชน',                    'type': 'str'},
    3:  {'field': 'dept',     'title': 'รหัสหน่วย',                         'type': 'dept'},
    4:  {'field': 'deptname', 'title': 'ชื่อหน่วย',                         'type': 'str'},
    5:  {'field': 'status',   'title': 'สถานะ',                            'type': 'str'},
    6:  {'field': 'name',     'title': 'ยศ-ชื่อ-สกุล',                      'type': 'str'},
    7:  {'field': 'addr',     'title': 'ที่อยู่',                            'type': 'str'},
    8:  {'field': 'tel',      'title': 'เบอร์โทร',                          'type': 'str'},
    9:  {'field': 'tmb',      'title': 'เลขบัญชี TTB',                     'type': 'str'},

    # --- ข้อมูลหุ้นและเงินปันผลเฉลี่ยคืน ---
    # หมายเหตุ: ในไฟล์ data-mast.xlsx ลำดับค่าจริงในตารางคือ:
    # Col 10: ทุนเรือนหุ้น (b1)
    # Col 11: ปันผล (p1)
    # Col 12: เฉลี่ยคืน (p2)
    # Col 13: รวมปันผล+เฉลี่ยคืน (sum_p)
    10: {'field': 'b1',       'title': 'ทุนเรือนหุ้นสะสม',                   'type': 'num'},
    11: {'field': 'p1',       'title': 'เงินปันผล',                         'type': 'num'},
    12: {'field': 'p2',       'title': 'เงินเฉลี่ยคืน',                      'type': 'num'},
    13: {'field': 'sum_p',    'title': 'รวมรับเงินปันผลเฉลี่ยคืน',            'type': 'num'},

    # --- หนี้สินคงเหลือ ---
    14: {'field': 'b2',       'title': 'เงินกู้สามัญคงเหลือ',                'type': 'num'},
    15: {'field': 'b3',       'title': 'เงินกู้พิเศษ1คงเหลือ',               'type': 'num'},
    16: {'field': 'b4',       'title': 'เงินกู้พิเศษ2คงเหลือ',               'type': 'num'},
    17: {'field': 'b5',       'title': 'เงินกู้พิเศษ3คงเหลือ',               'type': 'num'},
    18: {'field': 'b6',       'title': 'เงินกู้พิเศษ4คงเหลือ',               'type': 'num'},
    19: {'field': 'b71',      'title': 'ภาระค้ำประกันคงเหลือ 1',             'type': 'num'},
    20: {'field': 'b72',      'title': 'ภาระค้ำประกันคงเหลือ 2',             'type': 'num'},

    # --- รายการหักเงิน ---
    21: {'field': 'd1',       'title': 'หัก เงินประกันชีวิต กรมธรรม์ที่ 3',   'type': 'num'},
    22: {'field': 'd2',       'title': 'หัก เงินประกันชีวิต เงินกู้สามัญ',     'type': 'num'},
    23: {'field': 'd3',       'title': 'หัก เงินประกันภัย ผู้ค้ำประกัน',       'type': 'num'},
    24: {'field': 'd4',       'title': 'หัก เงินประกันชีวิต เงินกู้พิเศษ',      'type': 'num'},
    25: {'field': 'd5',       'title': 'หัก ชำระหนี้ ภาระค้ำประกัน',          'type': 'num'},
    26: {'field': 'd6',       'title': 'หัก ปันผล (ชำระต้นเงินสามัญ)',        'type': 'num'},
    27: {'field': 'd7',       'title': 'หัก ชำระหนี้อื่นๆ',                 'type': 'num'},
    28: {'field': 'd81',      'title': 'หัก ฌาปนกิจ สสอท.',                'type': 'num'},
    29: {'field': 'd82',      'title': 'หัก ฌาปนกิจ สส.ชสอ.',               'type': 'num'},
    30: {'field': 'd83',      'title': 'หัก ฌาปนกิจ สส.สท.',                'type': 'num'},
    31: {'field': 'd84',      'title': 'หัก ฌาปนกิจ สส.สก.',                'type': 'num'},

    # --- ยอดรวมสรุป ---
    32: {'field': 't1',       'title': 'รวมหักเงินทั้งสิ้น',                  'type': 'num'},
    33: {'field': 't2',       'title': 'ยอดรับสุทธิ',                        'type': 'num'}
}


def clean_num(val):
    """แปลงค่าตัวเลขให้เป็น int หรือ float ถ้าว่างให้คืนค่า 0"""
    if val is None or val == '':
        return 0
    try:
        f = float(val)
        return int(f) if f.is_integer() else f
    except (ValueError, TypeError):
        return 0


def clean_str(val):
    """แปลงค่าข้อความเป็นสตริงพร้อมตัดช่องว่าง"""
    if val is None:
        return ''
    return str(val).strip()


def extract_year_from_sheet_name(sheet_name):
    """
    ดึงปีจากชื่อ Sheet เช่น:
    - 'data2026' -> '2026'
    - '2026'     -> '2026'
    - 'data2569' -> '2026' (แปลง พ.ศ. 2569 เป็น ค.ศ. 2026)
    - '2568'     -> '2025'
    """
    clean_name = sheet_name.strip()
    digits = re.findall(r'\d{4}', clean_name)
    if digits:
        y_int = int(digits[0])
        # หากเป็นปี พ.ศ. (ตั้งแต่ 2500 ขึ้นไป) ให้แปลงเป็น ค.ศ.
        if y_int >= 2500:
            y_int -= 543
        return str(y_int)
    
    # กรณีไม่มีตัวเลข 4 หลัก ให้ตัดเฉพาะตัวอักษรและตัวเลข
    slug = re.sub(r'[^a-zA-Z0-9]', '', clean_name)
    return slug if slug else clean_name


def load_idp_lookup(base_dir):
    """โหลดข้อมูลเลขบัตรประชาชน (idp) จากไฟล์อ้างอิงเดิม (ds/data2025.json หรือ data2025.json) มาเตรียมไว้จับคู่"""
    lookup = {}
    data2025_json = os.path.join(base_dir, 'ds', 'data2025.json')
    if not os.path.exists(data2025_json):
        data2025_json = os.path.join(base_dir, 'data2025.json')
    if os.path.exists(data2025_json):
        try:
            with open(data2025_json, 'r', encoding='utf-8') as f:
                d25 = json.load(f)
                for item in d25:
                    m = clean_str(item.get('memo', '')).zfill(5)
                    idp_val = clean_str(item.get('idp', ''))
                    if m and idp_val:
                        lookup[m] = idp_val
            print(f" Loaded {len(lookup)} IDP lookups from {os.path.basename(data2025_json)}")
        except Exception as e:
            print("Notice: Could not load 2025 IDP reference:", e)
    return lookup


def process_sheet(ws, sheet_name, idp_lookup, base_dir):
    """ประมวลผล Sheet หนึ่งของ data-mast.xlsx และสร้างไฟล์ data<year>.json และ data<year>.xlsx"""
    year_key = extract_year_from_sheet_name(sheet_name)
    print(f"\n==================================================")
    print(f" Processing Sheet: '{sheet_name}' -> Year: '{year_key}'")
    print(f"==================================================")

    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        print(f"[WARN] Sheet '{sheet_name}' is empty or has only headers. Skipping.")
        return None

    headers = rows[0]
    data_rows = [r for r in rows[1:] if any(r is not None and str(r).strip() != '' for r in r)]
    print(f"Found {len(data_rows)} data rows.")

    records = []
    for r_idx, r in enumerate(data_rows, 1):
        # ข้ามแถวที่ไม่มีข้อมูลรหัสสมาชิก (Col 1)
        if len(r) < 2 or r[1] is None or str(r[1]).strip() == '':
            continue

        memo_str = clean_str(r[1]).zfill(5)

        # แปลงข้อมูลตาม COLUMN_TEMPLATE
        row_dict = {}
        for col_idx, meta in COLUMN_TEMPLATE.items():
            field_name = meta['field']
            val = r[col_idx] if col_idx < len(r) else None

            if meta['type'] == 'num':
                row_dict[field_name] = clean_num(val)
            elif meta['type'] == 'memo':
                row_dict[field_name] = memo_str
            elif meta['type'] == 'int':
                row_dict[field_name] = clean_num(val) or r_idx
            elif meta['type'] == 'dept':
                dept_str = clean_str(val)
                try:
                    row_dict[field_name] = int(dept_str)
                except ValueError:
                    row_dict[field_name] = dept_str
            else: # 'str'
                row_dict[field_name] = clean_str(val)

        # หากไม่มีเลขบัตรประชาชน ให้ค้นจาก IDP Lookup
        if not row_dict.get('idp'):
            row_dict['idp'] = idp_lookup.get(memo_str, '')

        # คำนวณยอดรวมตรวจสอบอัตโนมัติ (Sanity Checks)
        # 1. ยอดรวมเงินปันผลเฉลี่ยคืน sum_p = p1 + p2
        if not row_dict.get('sum_p'):
            row_dict['sum_p'] = row_dict.get('p1', 0) + row_dict.get('p2', 0)

        # 2. ยอดรับสุทธิ t2 = sum_p - t1
        if not row_dict.get('t2') and row_dict.get('sum_p'):
            row_dict['t2'] = row_dict.get('sum_p', 0) - row_dict.get('t1', 0)

        records.append(row_dict)

    # โฟลเดอร์จัดเก็บข้อมูล ds
    ds_dir = os.path.join(base_dir, 'ds')
    os.makedirs(ds_dir, exist_ok=True)

    # ส่งออกเป็น data<year>.json ในโฟลเดอร์ ds
    output_json = os.path.join(ds_dir, f"data{year_key}.json")
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    print(f"[OK] Successfully exported {len(records)} records to {output_json}")

    # ส่งออกเป็น data<year>.xlsx ในโฟลเดอร์ ds สำหรับดาวน์โหลด
    output_xlsx = os.path.join(ds_dir, f"data{year_key}.xlsx")
    export_sheet_to_excel(ws, headers, data_rows, output_xlsx)
    print(f"[OK] Successfully exported Excel to {output_xlsx}")

    return {
        'year': year_key,
        'sheet_name': sheet_name,
        'total': len(records),
        'json_file': output_json,
        'xlsx_file': output_xlsx
    }


def export_sheet_to_excel(src_ws, headers, data_rows, target_file):
    """ส่งออกข้อมูล Sheet เป็นไฟล์ Excel แยกรายปี"""
    new_wb = openpyxl.Workbook()
    dest_ws = new_wb.active
    dest_ws.title = "ข้อมูลสมาชิก"

    # เขียนหัวตาราง
    dest_ws.append(list(headers))

    # เขียนข้อมูลแถว
    for row in data_rows:
        dest_ws.append(list(row))

    new_wb.save(target_file)


def update_config_for_year(year_key, base_dir):
    """ตรวจสอบและสร้างค่า config เริ่มต้นสำหรับปีใหม่ใน config.json"""
    config_file = os.path.join(base_dir, 'config.json')
    if not os.path.exists(config_file):
        return

    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            cfg = json.load(f)

        if str(year_key) not in cfg:
            # คำนวณ พ.ศ.
            try:
                be_year = str(int(year_key) + 543)
                short_be = be_year[-2:]
            except:
                be_year = str(year_key)
                short_be = "69"

            cfg[str(year_key)] = {
                "year": be_year,
                "dividend_rate": "5.02%",
                "refund_rate": "13.50%",
                "balance_date": f"ณ 30 ก.ย.{short_be}",
                "meeting_title": f"**ประชุมใหญ่สามัญประจำปี {be_year}**",
                "meeting_date": f"วันพุธที่ 18 พฤศจิกายน {be_year}",
                "meeting_time_start": "เวลา 08.00- 14.30 น.",
                "meeting_vote_time": "เวลา 10.30 - 14.30 น.",
                "org_title": "ใบรายงานเงินปันผลเฉลี่ยคืนประจำปี",
                "org_name": "สหกรณ์ออมทรัพย์กองบัญชาการกองทัพไทย จำกัด",
                "post_box": "ตู้ ปณ.62 ปณฝ.ศูนย์ราชการเฉลิมพระเกียรติ 10215",
                "bank_name": "ธนาคารทหารไทยธนชาต จำกัด (มหาชน)",
                "sign_rank": "นาวาอากาศโท",
                "sign_name": "(สุวัฒน์ชัย  เรื่อศรีจันทร์)",
                "sign_position": "กรรมการดำเนินการ ทำหน้าที่ผู้จัดการ",
                "sign_image": "assets/img/sign2025.png"
            }
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            print(f" Added default config for year '{year_key}' ({be_year}) into config.json")
    except Exception as e:
        print("Notice: Error updating config.json:", e)


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))

    # 1. ตรวจสอบไฟล์ Excel: ใช้ data-mast.xlsx ใน path เดียวกันเป็นหลัก (Standalone / Portable)
    target_mast = None
    if len(sys.argv) > 1 and sys.argv[1].endswith(('.xlsx', '.xlsm')):
        custom_path = os.path.abspath(sys.argv[1])
        if os.path.exists(custom_path):
            target_mast = custom_path
            print(f"Using specified Excel file: {target_mast}")

    if not target_mast:
        local_mast = os.path.join(base_dir, 'data-mast.xlsx')
        if os.path.exists(local_mast):
            target_mast = local_mast
            print(f"Using primary data-mast.xlsx in same directory: {target_mast}")
        else:
            # Fallback สำรอง: ตรวจสอบโฟลเดอร์พาเรนต์ (กรณีเพิ่งติดตั้ง)
            root_dir = os.path.abspath(os.path.join(base_dir, '..'))
            root_mast = os.path.join(root_dir, 'data-mast.xlsx')
            if os.path.exists(root_mast):
                print(f"Notice: Local data-mast.xlsx not found. Copying from parent: {root_mast}")
                shutil.copy2(root_mast, local_mast)
                target_mast = local_mast
                print(f"Copied to local directory: {local_mast}")
            else:
                print(f"[ERROR] ไม่พบไฟล์ data-mast.xlsx ในโฟลเดอร์ {base_dir}")
                print(f"กรุณาวางไฟล์ data-mast.xlsx ไว้ในโฟลเดอร์เดียวกับ convert_data.py")
                return

    # โหลด IDP Lookup
    idp_lookup = load_idp_lookup(base_dir)

    # โหลด Workbook
    print(f"\nLoading workbook: {target_mast} ...")
    wb = openpyxl.load_workbook(target_mast, data_only=True)
    sheet_names = wb.sheetnames
    print(f"Available Sheets in data-mast.xlsx: {sheet_names}")

    results = []
    for s_name in sheet_names:
        ws = wb[s_name]
        res = process_sheet(ws, s_name, idp_lookup, base_dir)
        if res:
            results.append(res)
            update_config_for_year(res['year'], base_dir)

    print("\n==================================================")
    print(" SUMMARY OF DATA CONVERSION")
    print("==================================================")
    for r in results:
        print(f"• Sheet: '{r['sheet_name']}' -> Year '{r['year']}': {r['total']} records processed.")
    print("==================================================\n")

    # แจ้งเตือนเซิร์ฟเวอร์ผ่าน HTTP GET เพื่อ reload ข้อมูลทันที
    print("Sending reload notification to server (HTTP GET /api/reload)...")
    try:
        import urllib.request
        req = urllib.request.Request('http://localhost:8080/api/reload', headers={'User-Agent': 'convert_data'})
        with urllib.request.urlopen(req, timeout=2) as response:
            if response.status == 200:
                body = response.read().decode('utf-8')
                print(f"[OK] Server reloaded successfully: {body}")
    except Exception as e:
        print("Notice: Server is not currently running or reachable (data will be loaded when server starts).")


if __name__ == '__main__':
    main()
