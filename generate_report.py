import calendar
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import numpy as np


def build_monthly_report(raw_df: pd.DataFrame, target_year: int, target_month: int, output_path: str):
    """
    ฟังก์ชันแปลงข้อมูลราคาน้ำมันดิบให้ออกมาเป็นรายงาน Excel 3 ชีท 
    โดยจัดเรียงลำดับหัวข้อประเภทน้ำมันตามเงื่อนไขที่กำหนดอย่างแม่นยำ
    """
    # ----------------- 1. เตรียมตัวแปรวันที่และปฏิทิน -----------------
    target_dt = datetime(target_year, target_month, 1)
    prev_dt = target_dt - pd.DateOffset(months=1)
    
    num_days_curr = calendar.monthrange(target_year, target_month)[1]
    num_days_prev = calendar.monthrange(prev_dt.year, prev_dt.month)[1]
    
    days_cols = [str(d) for d in range(1, 32)]
    
    month_th = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
                "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
    month_th_abbr = ["", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", 
                     "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
    
    title_text = f"ราคาน้ำมัน {month_th[target_month]} {target_year + 543}"
    col_avg_prev = f"avg {month_th_abbr[prev_dt.month]} {(prev_dt.year + 543) % 100}"
    col_avg_curr = f"avg {month_th_abbr[target_month]} {(target_year + 543) % 100}"
    
    df = raw_df.copy()
    df['Date'] = pd.to_datetime(df['Date'])
    
    # ลำดับปั๊มมาตรฐานตามฟอร์ม
    standard_brands = [
        'ปตท (PTT)', 'เชลล์ (SHELL)', 'เอสโซ่ (ESSO)', 'เชฟรอน (CALTEX)', 
        'ทีพีไอ (TPI)', 'บางจาก (BCP)', 'ภาคใต้เชื้อเพลิง (PT)', 
        'สยามสหบริการ (SUSCO)', 'ทั่วไป', 'คอสโม (COSMO)', 
        'เพียว (PURE)', 'ปิโตรนาส (PETRONAS)'
    ]

    # ----------------- 2. กำหนดลำดับและชื่อหัวข้อน้ำมัน -----------------
    # กำหนด Keyword สำหรับ Match กับชื่อน้ำมันในไฟล์ดิบ และชื่อหัวข้อเป๊ะๆ ที่ต้องการแสดง
    priority_oil_configs = [
        {
            "priority": 1,
            "title": "1. เบนซิน ออกเทน 95 (ULG 95 RON)  (5220003)",
            "match": lambda name: ('เบนซิน' in name or 'ULG' in name.upper()) and ('95' in name) and ('แก๊ส' not in name and 'GAS' not in name.upper())
        },
        {
            "priority": 2,
            "title": "2. ดีเซลหมุนเร็ว (HSD, 0.005%S)  (5220004)",
            "match": lambda name: ('ดีเซล' in name or 'HSD' in name.upper()) and ('B10' not in name and 'B20' not in name and 'B7' not in name)
        },
        {
            "priority": 3,
            "title": "3. แก๊สโซฮอล ออกเทน 91 (Gasohol 91-E10)  (5220006)",
            "match": lambda name: ('91' in name) and ('แก๊ส' in name or 'GAS' in name.upper())
        },
        {
            "priority": 4,
            "title": "4. แก๊สโซฮอล ออกเทน 95 (Gasohol 95-E10)  (5220007)",
            "match": lambda name: ('95' in name) and ('แก๊ส' in name or 'GAS' in name.upper()) and ('E20' not in name.upper() and 'E85' not in name.upper())
        },
        {
            "priority": 5,
            "title": "5. แก๊สโซฮอล E20 เบนซินไร้สารตะกั่ว 80% + แอลกอฮอล์ 20% (Gasohol 95-E20) (5220010)",
            "match": lambda name: 'E20' in name.upper()
        },
        {
            "priority": 6,
            "title": "6. แก๊สโซฮอล E85 ออกเทน 95 (Gasohol 95-E85) (5220011)",
            "match": lambda name: 'E85' in name.upper()
        },
        {
            "priority": 7,
            "title": "7. ดีเซลหมุนเร็ว (B10) (HSD, 0.005%S)  (5220012)",
            "match": lambda name: 'B10' in name.upper()
        }
    ]

    def match_product(prod_name):
        for cfg in priority_oil_configs:
            if cfg["match"](prod_name):
                return cfg["priority"], cfg["title"]
        return 999, prod_name

    # ดึงชื่อผลิตภัณฑ์ทั้งหมดที่มีในฐานข้อมูล
    all_raw_products = [p for p in df['Product'].dropna().unique() if str(p).strip() != '']
    
    # จัดกลุ่มและเรียงลำดับรายการน้ำมัน
    sorted_products = []
    matched_flags = {}
    
    for cfg in priority_oil_configs:
        for p in all_raw_products:
            if p not in matched_flags and cfg["match"](p):
                sorted_products.append((p, cfg["title"], cfg["priority"]))
                matched_flags[p] = True
                break
                
    # นำน้ำมันประเภทอื่นๆ ที่เหลือมาต่อท้าย
    other_priority = 8
    for p in all_raw_products:
        if p not in matched_flags:
            sorted_products.append((p, f"{other_priority}. {p}", other_priority))
            other_priority += 1

    # ----------------- 3. ฟังก์ชันเตรียมข้อมูลตาราง -----------------
    def prepare_sheet_data(price_offset=0.0):
        rows = []
        
        for raw_prod_name, display_title, _ in sorted_products:
            prod_df = df[df['Product'] == raw_prod_name].copy()
            if price_offset > 0:
                prod_df['Price'] = prod_df['Price'] + price_offset

            curr_df = prod_df[(prod_df['Date'].dt.year == target_year) & (prod_df['Date'].dt.month == target_month)]
            prev_df = prod_df[(prod_df['Date'].dt.year == prev_dt.year) & (prod_df['Date'].dt.month == prev_dt.month)]
            
            prev_avg = prev_df[prev_df['Price'] > 0].groupby('Brand')['Price'].mean()
            prev_last_day_val = prev_df[prev_df['Date'].dt.day == num_days_prev].set_index('Brand')['Price']
            
            if not curr_df.empty:
                pivot_curr = curr_df.pivot_table(index='Brand', columns=curr_df['Date'].dt.day, values='Price', aggfunc='last')
            else:
                pivot_curr = pd.DataFrame()

            available_brands = [b for b in standard_brands if b in prod_df['Brand'].unique()]
            if not available_brands:
                available_brands = list(prod_df['Brand'].unique())

            # เพิ่มแถว Header หมวดหมู่ประเภทน้ำมัน
            rows.append({
                'is_header': True,
                'title': display_title,
                'data': None
            })
            
            for brand in available_brands:
                r = {'Brand': brand}
                p_avg = prev_avg.get(brand, np.nan)
                r['avg_prev'] = p_avg
                r['prev_31'] = prev_last_day_val.get(brand, np.nan)
                
                valid_daily_prices = []
                for d in range(1, 32):
                    if d <= num_days_curr and (brand in pivot_curr.index) and (d in pivot_curr.columns):
                        val = pivot_curr.loc[brand, d]
                        if pd.notna(val) and val > 0:
                            r[str(d)] = float(val)
                            valid_daily_prices.append(float(val))
                        else:
                            r[str(d)] = 0.00
                    else:
                        r[str(d)] = 0.00
                
                c_avg = np.mean(valid_daily_prices) if valid_daily_prices else 0.00
                r['avg_curr'] = c_avg
                
                if pd.notna(p_avg) and p_avg > 0 and c_avg > 0:
                    r['REL'] = (c_avg / p_avg) * 100.0
                else:
                    r['REL'] = 0.00
                    
                rows.append({
                    'is_header': False,
                    'title': '',
                    'data': r
                })
        return rows

    # ----------------- 4. จัดรูปแบบและสร้างไฟล์ Excel -----------------
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    font_title = Font(name='Angsana New', size=16, bold=True)
    font_bold = Font(name='Angsana New', size=13, bold=True)
    font_regular = Font(name='Angsana New', size=12)
    
    fill_blue_header = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    fill_sub_green = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='A6A6A6'),
        right=Side(style='thin', color='A6A6A6'),
        top=Side(style='thin', color='A6A6A6'),
        bottom=Side(style='thin', color='A6A6A6')
    )
    align_center = Alignment(horizontal='center', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')

    # 4.1 Sheet 1 และ Sheet 2
    sheets_config = [
        ('ราคาน้ำมันขายปลีก', 0.0),
        ('ราคารวมหัวจ่าย (+0.05)', 0.05)
    ]
    
    headers_main = ['ปั๊มน้ำมัน/วันที่', col_avg_prev, '31'] + days_cols + [col_avg_curr, 'REL']

    for sheet_name, offset in sheets_config:
        ws = wb.create_sheet(title=sheet_name)
        matrix = prepare_sheet_data(price_offset=offset)
        
        ws.cell(row=1, column=1, value=title_text).font = font_title
        
        for col_idx, h_text in enumerate(headers_main, 1):
            cell = ws.cell(row=3, column=col_idx, value=h_text)
            cell.font = font_bold
            cell.alignment = align_center
            cell.border = thin_border
            if col_idx in [2, len(headers_main) - 1]:
                cell.fill = fill_yellow
            elif col_idx == 3:
                cell.fill = fill_sub_green

        current_row = 4
        for item in matrix:
            if item['is_header']:
                c = ws.cell(row=current_row, column=1, value=item['title'])
                c.font = font_bold
                c.alignment = align_left
                for col_idx in range(1, len(headers_main) + 1):
                    ws.cell(row=current_row, column=col_idx).fill = fill_blue_header
                    ws.cell(row=current_row, column=col_idx).border = thin_border
                current_row += 1
            else:
                d = item['data']
                b_cell = ws.cell(row=current_row, column=1, value=d['Brand'])
                b_cell.font = font_regular
                b_cell.alignment = align_left
                b_cell.border = thin_border
                
                p_cell = ws.cell(row=current_row, column=2, value=d['avg_prev'])
                p_cell.number_format = '0.00'
                p_cell.font = font_regular
                p_cell.fill = fill_yellow
                p_cell.alignment = align_right
                p_cell.border = thin_border
                
                prev31_cell = ws.cell(row=current_row, column=3, value=d['prev_31'])
                prev31_cell.number_format = '0.00'
                prev31_cell.font = font_regular
                prev31_cell.fill = fill_sub_green
                prev31_cell.alignment = align_right
                prev31_cell.border = thin_border
                
                for idx, day_str in enumerate(days_cols, 4):
                    day_cell = ws.cell(row=current_row, column=idx, value=d[day_str])
                    day_cell.number_format = '0.00'
                    day_cell.font = font_regular
                    day_cell.alignment = align_right
                    day_cell.border = thin_border
                
                c_avg_cell = ws.cell(row=current_row, column=len(headers_main) - 1, value=d['avg_curr'])
                c_avg_cell.number_format = '0.00'
                c_avg_cell.font = font_regular
                c_avg_cell.fill = fill_yellow
                c_avg_cell.alignment = align_right
                c_avg_cell.border = thin_border
                
                rel_cell = ws.cell(row=current_row, column=len(headers_main), value=d['REL'])
                rel_cell.number_format = '0.00'
                rel_cell.font = font_regular
                rel_cell.alignment = align_right
                rel_cell.border = thin_border
                
                current_row += 1

    # 4.2 Sheet 3 (ผลต่างราคารายวันเทียบวันก่อนหน้า)
    ws3 = wb.create_sheet(title='ผลต่างราคารายวัน')
    ws3.cell(row=1, column=1, value=f"{title_text} (ส่วนต่างราคาเทียบวันก่อนหน้า)").font = font_title
    
    headers_diff = ['ปั๊มน้ำมัน/วันที่', '31'] + days_cols + ['AI (Avg Diff)']
    
    for col_idx, h_text in enumerate(headers_diff, 1):
        cell = ws3.cell(row=3, column=col_idx, value=h_text)
        cell.font = font_bold
        cell.alignment = align_center
        cell.border = thin_border
        if col_idx == 2:
            cell.fill = fill_sub_green
        elif col_idx == len(headers_diff):
            cell.fill = fill_yellow

    base_data = prepare_sheet_data(price_offset=0.0)
    current_row = 4
    
    for item in base_data:
        if item['is_header']:
            c = ws3.cell(row=current_row, column=1, value=item['title'])
            c.font = font_bold
            c.alignment = align_left
            for col_idx in range(1, len(headers_diff) + 1):
                ws3.cell(row=current_row, column=col_idx).fill = fill_blue_header
                ws3.cell(row=current_row, column=col_idx).border = thin_border
            current_row += 1
        else:
            d = item['data']
            b_cell = ws3.cell(row=current_row, column=1, value=d['Brand'])
            b_cell.font = font_regular
            b_cell.alignment = align_left
            b_cell.border = thin_border
            
            ws3.cell(row=current_row, column=2, value=0.00).number_format = '0.00'
            ws3.cell(row=current_row, column=2).fill = fill_sub_green
            ws3.cell(row=current_row, column=2).border = thin_border
            ws3.cell(row=current_row, column=2).alignment = align_right
            
            diff_values = []
            prev_price = d['prev_31'] if pd.notna(d['prev_31']) else 0.00
            
            for idx, day_str in enumerate(days_cols, 3):
                curr_price = d[day_str]
                
                if curr_price > 0 and prev_price > 0:
                    diff = round(curr_price - prev_price, 2)
                    diff_values.append(diff)
                else:
                    diff = 0.00
                
                cell_diff = ws3.cell(row=current_row, column=idx, value=diff)
                cell_diff.number_format = '0.00'
                cell_diff.font = font_regular
                cell_diff.alignment = align_right
                cell_diff.border = thin_border
                
                if curr_price > 0:
                    prev_price = curr_price
            
            avg_diff = np.mean(diff_values) if diff_values else 0.00
            ai_cell = ws3.cell(row=current_row, column=len(headers_diff), value=avg_diff)
            ai_cell.number_format = '0.00'
            ai_cell.font = font_bold
            ai_cell.fill = fill_yellow
            ai_cell.alignment = align_right
            ai_cell.border = thin_border
            
            current_row += 1

    # ปรับขนาดความกว้างคอลัมน์
    for sheet in wb.worksheets:
        sheet.column_dimensions['A'].width = 24
        for col_letter in [get_column_letter(c) for c in range(2, sheet.max_column + 1)]:
            sheet.column_dimensions[col_letter].width = 7.5

    wb.save(output_path)
    print(f"บันทึกไฟล์รายงานสำเร็จเรียบร้อยที่: {output_path}")
