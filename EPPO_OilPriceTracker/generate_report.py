import calendar
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import numpy as np


def build_monthly_report(raw_df: pd.DataFrame, target_year: int, target_month: int, output_path: str):
    """
    ฟังก์ชันแปลงข้อมูลราคาน้ำมันดิบให้ออกมาเป็นรายงาน Excel 3 ชีท ตามโครงสร้างและสไตล์ที่กำหนด
    """
    # ----------------- 1. เตรียมตัวแปรวันที่และปฏิทิน -----------------
    target_dt = datetime(target_year, target_month, 1)
    prev_dt = target_dt - pd.DateOffset(months=1)
    
    # จำนวนวันของเดือนปัจจุบันและเดือนก่อนหน้า
    num_days_curr = calendar.monthrange(target_year, target_month)[1]
    num_days_prev = calendar.monthrange(prev_dt.year, prev_dt.month)[1]
    
    # กำหนดคอลัมน์วันที่ 1 ถึง 31 ของเดือนปัจจุบัน
    days_cols = [str(d) for d in range(1, 32)]
    
    # ชื่อเดือนภาษาไทยแบบเต็มและย่อ
    month_th = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
                "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
    month_th_abbr = ["", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", 
                     "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
    
    title_text = f"ราคาน้ำมัน {month_th[target_month]} {target_year + 543}"
    col_avg_prev = f"avg {month_th_abbr[prev_dt.month]} {(prev_dt.year + 543) % 100}"
    col_avg_curr = f"avg {month_th_abbr[target_month]} {(target_year + 543) % 100}"
    
    # แปลงคอลัมน์ Date เป็น Datetime
    df = raw_df.copy()
    df['Date'] = pd.to_datetime(df['Date'])
    
    # รายชื่อแบรนด์มาตรฐานตามโครงสร้างในรายงาน (หรือดึงตามลำดับที่มีในข้อมูล)
    standard_brands = [
        'ปตท (PTT)', 'เชลล์ (SHELL)', 'เอสโซ่ (ESSO)', 'เชฟรอน (CALTEX)', 
        'ทีพีไอ (TPI)', 'บางจาก (BCP)', 'ภาคใต้เชื้อเพลิง (PT)', 
        'สยามสหบริการ (SUSCO)', 'ทั่วไป', 'คอสโม (COSMO)', 
        'เพียว (PURE)', 'ปิโตรนาส (PETRONAS)'
    ]
    
    # ----------------- 2. ฟังก์ชันจัดโครงสร้างตารางข้อมูล -----------------
    def prepare_sheet_data(price_offset=0.0):
        rows = []
        products = df['Product'].dropna().unique()
        
        for prod_idx, prod in enumerate(products, 1):
            prod_df = df[df['Product'] == prod].copy()
            if price_offset > 0:
                prod_df['Price'] = prod_df['Price'] + price_offset

            # แยกชุดข้อมูลเดือนก่อนหน้า และเดือนปัจจุบัน
            curr_df = prod_df[(prod_df['Date'].dt.year == target_year) & (prod_df['Date'].dt.month == target_month)]
            prev_df = prod_df[(prod_df['Date'].dt.year == prev_dt.year) & (prod_df['Date'].dt.month == prev_dt.month)]
            
            # คำนวณ avg เดือนก่อนหน้า (เฉลี่ยเฉพาะวันที่มีราคาขายจริง)
            prev_avg = prev_df[prev_df['Price'] > 0].groupby('Brand')['Price'].mean()
            
            # ราคาของวันสุดท้ายเดือนก่อนหน้า (วันที่ 28/29/30/31)
            prev_last_day_val = prev_df[prev_df['Date'].dt.day == num_days_prev].set_index('Brand')['Price']
            
            # Pivot ราคารายวันของเดือนปัจจุบัน
            if not curr_df.empty:
                pivot_curr = curr_df.pivot_table(index='Brand', columns=curr_df['Date'].dt.day, values='Price', aggfunc='last')
            else:
                pivot_curr = pd.DataFrame()

            # รวมแบรนด์ที่มีในข้อมูล
            available_brands = [b for b in standard_brands if b in prod_df['Brand'].unique()]
            if not available_brands:
                available_brands = list(prod_df['Brand'].unique())

            # เพิ่มหัวหมวดหมู่น้ำมัน
            rows.append({
                'is_header': True,
                'title': f"{prod_idx}. {prod}",
                'data': None
            })
            
            for brand in available_brands:
                r = {'Brand': brand}
                p_avg = prev_avg.get(brand, np.nan)
                r['avg_prev'] = p_avg
                r['prev_31'] = prev_last_day_val.get(brand, np.nan)
                
                # รายวัน 1 - 31 (หากเดือนมีไม่ถึง 31 วัน หรือไม่มีข้อมูล ให้ใส่ 0.00)
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
                
                # คำนวณค่าเฉลี่ยเดือนปัจจุบัน (เฉพาะวันที่มีการขายจริง > 0)
                c_avg = np.mean(valid_daily_prices) if valid_daily_prices else 0.00
                r['avg_curr'] = c_avg
                
                # อัตราการเปลี่ยนแปลง Relative Change (REL) = (avg_curr / avg_prev) * 100
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

    # ----------------- 3. การสร้างไฟล์และการจัดสไตล์ Excel -----------------
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # เอาชีทเปล่าเริ่มต้นออก
    
    # ฟอนต์และสีตามแบบฟอร์ม
    font_title = Font(name='Angsana New', size=16, bold=True)
    font_bold = Font(name='Angsana New', size=13, bold=True)
    font_regular = Font(name='Angsana New', size=12)
    
    fill_blue_header = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    fill_sub_green = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    fill_red_alert = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    font_white = Font(name='Angsana New', size=12, color="FFFFFF", bold=True)
    
    thin_border = Border(
        left=Side(style='thin', color='A6A6A6'),
        right=Side(style='thin', color='A6A6A6'),
        top=Side(style='thin', color='A6A6A6'),
        bottom=Side(style='thin', color='A6A6A6')
    )
    align_center = Alignment(horizontal='center', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')

    # 3.1 สร้าง Sheet 1 (ราคาปกติ) และ Sheet 2 (ราคาบวกหัวจ่าย +0.05)
    sheets_config = [
        ('ราคาน้ำมันขายปลีก', 0.0),
        ('ราคารวมหัวจ่าย (+0.05)', 0.05)
    ]
    
    headers_main = ['ปั๊มน้ำมัน/วันที่', col_avg_prev, '31'] + days_cols + [col_avg_curr, 'REL']

    for sheet_name, offset in sheets_config:
        ws = wb.create_sheet(title=sheet_name)
        matrix = prepare_sheet_data(price_offset=offset)
        
        # เขียน Title
        ws.cell(row=1, column=1, value=title_text).font = font_title
        
        # เขียน Header แถวที่ 3
        for col_idx, h_text in enumerate(headers_main, 1):
            cell = ws.cell(row=3, column=col_idx, value=h_text)
            cell.font = font_bold
            cell.alignment = align_center
            cell.border = thin_border
            if col_idx in [2, len(headers_main) - 1]:
                cell.fill = fill_yellow
            elif col_idx == 3:
                cell.fill = fill_sub_green
            else:
                cell.fill = fill_sub_green if col_idx <= 3 else PatternFill(fill_type=None)

        current_row = 4
        for item in matrix:
            if item['is_header']:
                # แถบสีฟ้าน้ำมันแต่ละประเภท
                c = ws.cell(row=current_row, column=1, value=item['title'])
                c.font = font_bold
                c.alignment = align_left
                for col_idx in range(1, len(headers_main) + 1):
                    ws.cell(row=current_row, column=col_idx).fill = fill_blue_header
                    ws.cell(row=current_row, column=col_idx).border = thin_border
                current_row += 1
            else:
                d = item['data']
                # คอลัมน์ 1: แบรนด์
                b_cell = ws.cell(row=current_row, column=1, value=d['Brand'])
                b_cell.font = font_regular
                b_cell.alignment = align_left
                b_cell.border = thin_border
                
                # คอลัมน์ 2: avg เดือนก่อนหน้า
                p_cell = ws.cell(row=current_row, column=2, value=d['avg_prev'])
                p_cell.number_format = '0.00'
                p_cell.font = font_regular
                p_cell.fill = fill_yellow
                p_cell.alignment = align_right
                p_cell.border = thin_border
                
                # คอลัมน์ 3: วันที่ 31 ของเดือนก่อน
                prev31_cell = ws.cell(row=current_row, column=3, value=d['prev_31'])
                prev31_cell.number_format = '0.00'
                prev31_cell.font = font_regular
                prev31_cell.fill = fill_sub_green
                prev31_cell.alignment = align_right
                prev31_cell.border = thin_border
                
                # คอลัมน์ 4-34: รายวัน 1 - 31
                for idx, day_str in enumerate(days_cols, 4):
                    day_cell = ws.cell(row=current_row, column=idx, value=d[day_str])
                    day_cell.number_format = '0.00'
                    day_cell.font = font_regular
                    day_cell.alignment = align_right
                    day_cell.border = thin_border
                
                # คอลัมน์ avg เดือนปัจจุบัน
                c_avg_cell = ws.cell(row=current_row, column=len(headers_main) - 1, value=d['avg_curr'])
                c_avg_cell.number_format = '0.00'
                c_avg_cell.font = font_regular
                c_avg_cell.fill = fill_yellow
                c_avg_cell.alignment = align_right
                c_avg_cell.border = thin_border
                
                # คอลัมน์ REL
                rel_cell = ws.cell(row=current_row, column=len(headers_main), value=d['REL'])
                rel_cell.number_format = '0.00'
                rel_cell.font = font_regular
                rel_cell.alignment = align_right
                rel_cell.border = thin_border
                
                current_row += 1

    # 3.2 สร้าง Sheet 3 (ผลต่างราคารายวันเทียบวันก่อนหน้า)
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
            # คอลัมน์ 1: แบรนด์
            b_cell = ws3.cell(row=current_row, column=1, value=d['Brand'])
            b_cell.font = font_regular
            b_cell.alignment = align_left
            b_cell.border = thin_border
            
            # คอลัมน์ 2: วันที่ 31 ของเดือนก่อน (ตั้งต้นความต่างเป็น 0.00)
            ws3.cell(row=current_row, column=2, value=0.00).number_format = '0.00'
            ws3.cell(row=current_row, column=2).fill = fill_sub_green
            ws3.cell(row=current_row, column=2).border = thin_border
            ws3.cell(row=current_row, column=2).alignment = align_right
            
            diff_values = []
            prev_price = d['prev_31'] if pd.notna(d['prev_31']) else 0.00
            
            # คำนวณความต่างรายวัน (วันปัจจุบัน - วันก่อนหน้า)
            for idx, day_str in enumerate(days_cols, 3):
                curr_price = d[day_str]
                
                # หากทั้งสองวันมีราคาขายจริง (> 0) ให้หาผลต่าง
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
                
                # เลื่อนราคาวันปัจจุบันไปเป็นวันก่อนหน้าของรอบถัดไป (ถ้าวันนั้นมีราคาขายจริง)
                if curr_price > 0:
                    prev_price = curr_price
            
            # คอลัมน์ AI (Average Diff ตลอดเดือน)
            avg_diff = np.mean(diff_values) if diff_values else 0.00
            ai_cell = ws3.cell(row=current_row, column=len(headers_diff), value=avg_diff)
            ai_cell.number_format = '0.00'
            ai_cell.font = font_bold
            ai_cell.fill = fill_yellow
            ai_cell.alignment = align_right
            ai_cell.border = thin_border
            
            current_row += 1

    # ปรับความกว้างคอลัมน์ให้อ่านง่ายในทุกชีท
    for sheet in wb.worksheets:
        sheet.column_dimensions['A'].width = 24
        for col_letter in [get_column_letter(c) for c in range(2, sheet.max_column + 1)]:
            sheet.column_dimensions[col_letter].width = 7.5

    wb.save(output_path)
    print(f"บันทึกไฟล์รายงานสำเร็จเรียบร้อยที่: {output_path}")