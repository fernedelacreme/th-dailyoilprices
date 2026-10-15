import os
from datetime import datetime, timedelta
import pandas as pd
import requests
from bs4 import BeautifulSoup
from io import BytesIO

# ------------------ ตั้งค่าไฟล์และตัวแปรเริ่มต้น ------------------
output_file = 'EPPO_Historical_Prices_Merged.xlsx'
url = "https://www2.eppo.go.th/epporop/entryreport/ropbydaypublic.aspx"
default_start_date = '2025-01-01'

today = datetime.today().date()
existing_df = None

# ใส่ Headers จำลองเป็น Google Chrome เพื่อป้องกันเว็บภาครัฐบล็อก IP Cloud
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
    'Accept-Language': 'th,en-US;q=0.9,en;q=0.8'
}

# ------------------ 1. ตรวจสอบไฟล์เดิมและหาวันที่ต้องดึงต่อ ------------------
if os.path.exists(output_file):
    print(f"กำลังอ่านข้อมูลจาก '{output_file}' เพื่อตรวจสอบวันที่ล่าสุด...")
    existing_df = pd.read_excel(output_file)
    existing_df['Date'] = pd.to_datetime(existing_df['Date'])
    
    if not existing_df.empty and existing_df['Date'].notna().any():
        last_date = existing_df['Date'].max().date()
        start_date = last_date + timedelta(days=1)
        print(f"วันที่ล่าสุดในไฟล์คือ: {last_date.strftime('%d/%m/%Y')}")
    else:
        start_date = pd.to_datetime(default_start_date).date()
else:
    print(f"ไม่พบไฟล์เดิม จะเริ่มดึงข้อมูลตั้งแต่เริ่มต้น ({default_start_date})...")
    start_date = pd.to_datetime(default_start_date).date()

if start_date > today:
    print(f"ข้อมูลในไฟล์เป็นปัจจุบันแล้ว (ถึงวันที่ {today.strftime('%d/%m/%Y')}) ไม่จำเป็นต้องดึงเพิ่ม")
    exit(0)

print(f"ช่วงวันที่ต้องดึงข้อมูล: {start_date.strftime('%d/%m/%Y')} ถึง {today.strftime('%d/%m/%Y')}\n")

date_list = pd.date_range(start=start_date, end=today).strftime('%d/%m/%Y')
all_data = []

# ------------------ 2. ดึงข้อมูลจากเว็บไซต์ EPPO ------------------
with requests.Session() as session:
    session.headers.update(headers)
    for date_str in date_list:
        print(f"กำลังดึงและแปลงข้อมูลวันที่: {date_str}")
        try:
            # 1. ยิง GET เพื่อดึง Token
            res_get = session.get(url, timeout=30)
            soup = BeautifulSoup(res_get.text, 'html.parser')
            
            viewstate = soup.find('input', {'name': '__VIEWSTATE'})
            viewstate_gen = soup.find('input', {'name': '__VIEWSTATEGENERATOR'})
            event_val = soup.find('input', {'name': '__EVENTVALIDATION'})
            
            if not viewstate or not event_val:
                print(f"  -> ไม่พบค่า ViewState/Token จากหน้าเว็บ EPPO สำหรับวันที่ {date_str} (อาจถูกจำกัดสิทธิ์ชั่วคราว)")
                continue

            payload = {
                '__EVENTTARGET': '',
                '__EVENTARGUMENT': '',
                '__VIEWSTATE': viewstate['value'],
                '__VIEWSTATEGENERATOR': viewstate_gen['value'] if viewstate_gen else '',
                '__EVENTVALIDATION': event_val['value'],
                'TbxToDate': date_str,
                'BtnGenerate.x': '50',
                'BtnGenerate.y': '10'
            }
            
            # 2. ยิง POST เพื่อขอไฟล์ Excel
            res_post = session.post(url, data=payload, timeout=30)
            
            # 3. อ่านไฟล์ Excel จาก Memory
            df_raw = pd.read_excel(BytesIO(res_post.content), header=None)
            
            brands = df_raw.iloc[11, 1:11].fillna('').astype(str).tolist()
            products = df_raw.iloc[13:25, 0].fillna('').astype(str).tolist()
            
            prices = df_raw.iloc[13:25, 1:11].copy()
            prices.columns = brands
            prices['Product'] = products
            prices['Date'] = date_str
            
            # 4. แปลงโครงสร้างแบบ Melt (Unpivot)
            df_melted = prices.melt(
                id_vars=['Date', 'Product'], 
                value_vars=brands, 
                var_name='Brand', 
                value_name='Price'
            )
            
            # 5. ดึง Effective Date
            eff_dates = df_raw.iloc[25, 1:11].fillna('').astype(str).tolist()
            eff_dict = dict(zip(brands, eff_dates))
            df_melted['Effective_Date'] = df_melted['Brand'].map(eff_dict)
            
            # ตัดค่าว่างทิ้ง
            df_melted = df_melted.dropna(subset=['Price'])
            
            all_data.append(df_melted)
            
        except Exception as e:
            print(f"  -> ข้ามวันที่ {date_str} เนื่องจาก: {e}")

# ------------------ 3. รวมข้อมูลและบันทึกผล ------------------
if all_data:
    new_df = pd.concat(all_data, ignore_index=True)
    new_df['Date'] = pd.to_datetime(new_df['Date'], format='%d/%m/%Y')

    if existing_df is not None and not existing_df.empty:
        print("\nกำลังนำข้อมูลใหม่ไปต่อท้ายข้อมูลเดิม...")
        master_df = pd.concat([existing_df, new_df], ignore_index=True)
        master_df.drop_duplicates(subset=['Date', 'Brand', 'Product'], keep='last', inplace=True)
    else:
        master_df = new_df

    master_df.sort_values(by=['Brand', 'Date', 'Product'], inplace=True)

    with pd.ExcelWriter(output_file, engine='openpyxl', datetime_format='DD/MM/YYYY') as writer:
        master_df.to_excel(writer, index=False)

    print(f"\nอัปเดตไฟล์สำเร็จ! รวมทั้งหมดมีข้อมูล {len(master_df):,} แถว")
else:
    print("\nไม่มีข้อมูลใหม่ถูกบันทึกเพิ่มเติม")
