import os
import time
from datetime import datetime, timedelta, timezone
import pandas as pd
import requests
from bs4 import BeautifulSoup
from io import BytesIO

# ------------------ ตั้งค่า ------------------
data_dir = 'data'
os.makedirs(data_dir, exist_ok=True)

url = "https://www2.eppo.go.th/epporop/entryreport/ropbydaypublic.aspx"

# กำหนด Timezone เวลาประเทศไทย (UTC+7)
tz_th = timezone(timedelta(hours=7))
today = datetime.now(tz_th).date()
current_year = today.year
output_file = os.path.join(data_dir, f'EPPO_Prices_{current_year}.parquet')

# Headers เสมือน Browser จริงแบบครบชุด เพื่อทะลุ Bot Shield ของภาครัฐ
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
    'Accept-Language': 'th-TH,th;q=0.9,en-US;q=0.8,en;q=0.7',
    'Accept-Encoding': 'gzip, deflate, br',
    'Connection': 'keep-alive',
    'Referer': 'https://www2.eppo.go.th/epporop/entryreport/ropbydaypublic.aspx',
    'Sec-Ch-Ua': '"Google Chrome";v="129", "Not=A?Brand";v="8", "Chromium";v="129"',
    'Sec-Ch-Ua-Mobile': '?0',
    'Sec-Ch-Ua-Platform': '"Windows"',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'same-origin',
    'Sec-Fetch-User': '?1',
    'Upgrade-Insecure-Requests': '1'
}

# ------------------ 1. ตรวจสอบข้อมูลเดิมของปีปัจจุบัน ------------------
existing_df = None
if os.path.exists(output_file):
    print(f"กำลังอ่านข้อมูลเดิมจาก '{output_file}'...")
    existing_df = pd.read_parquet(output_file)
    existing_df['Date'] = pd.to_datetime(existing_df['Date'])
    
    if not existing_df.empty and existing_df['Date'].notna().any():
        last_date = existing_df['Date'].max().date()
        start_date = last_date + timedelta(days=1)
        print(f"วันที่ล่าสุดในไฟล์คือ: {last_date.strftime('%d/%m/%Y')}")
    else:
        start_date = datetime(current_year, 1, 1).date()
else:
    print(f"ไม่พบไฟล์ปี {current_year} จะเริ่มดึงตั้งแต่วันที่ 01/01/{current_year}...")
    start_date = datetime(current_year, 1, 1).date()

if start_date > today:
    print(f"ข้อมูลปี {current_year} เป็นปัจจุบันแล้ว (ถึง {today.strftime('%d/%m/%Y')})")
    exit(0)

print(f"ช่วงวันที่ต้องดึงข้อมูล: {start_date.strftime('%d/%m/%Y')} ถึง {today.strftime('%d/%m/%Y')}\n")
date_list = pd.date_range(start=start_date, end=today).strftime('%d/%m/%Y')
all_data = []

# ------------------ 2. ดึงข้อมูลจาก EPPO ------------------
with requests.Session() as session:
    session.headers.update(headers)
    
    for date_str in date_list:
        print(f"กำลังดึงและแปลงข้อมูลวันที่: {date_str}")
        success_for_this_day = False
        
        # กลไก Retry สูงสุด 3 ครั้ง หากรอบแรกไม่เจอ ViewState
        for attempt in range(1, 4):
            try:
                res_get = session.get(url, timeout=30)
                soup = BeautifulSoup(res_get.text, 'html.parser')
                
                viewstate = soup.find('input', {'name': '__VIEWSTATE'})
                viewstate_gen = soup.find('input', {'name': '__VIEWSTATEGENERATOR'})
                event_val = soup.find('input', {'name': '__EVENTVALIDATION'})
                
                if not viewstate or not event_val:
                    print(f"  -> พยายามครั้งที่ {attempt}: ไม่พบ ViewState (สถานะตอบกลับ: {res_get.status_code}) รอ 3 วินาที...")
                    time.sleep(3)
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
                
                # ยิง POST เพื่อดาวน์โหลดตารางราคา
                res_post = session.post(url, data=payload, timeout=30)
                
                # ตรวจสอบว่าได้ไฟล์ Excel จริงหรือไม่
                if len(res_post.content) < 500:
                    print(f"  -> ข้อมูลที่ได้สั้นผิดปกติ อาจยังไม่มีรายงานสำหรับวันที่ {date_str}")
                    break
                    
                df_raw = pd.read_excel(BytesIO(res_post.content), header=None)
                
                brands = df_raw.iloc[11, 1:11].fillna('').astype(str).tolist()
                products = df_raw.iloc[13:25, 0].fillna('').astype(str).tolist()
                
                prices = df_raw.iloc[13:25, 1:11].copy()
                prices.columns = brands
                prices['Product'] = products
                prices['Date'] = date_str
                
                df_melted = prices.melt(
                    id_vars=['Date', 'Product'], 
                    value_vars=brands, 
                    var_name='Brand', 
                    value_name='Price'
                )
                
                eff_dates = df_raw.iloc[25, 1:11].fillna('').astype(str).tolist()
                eff_dict = dict(zip(brands, eff_dates))
                df_melted['Effective_Date'] = df_melted['Brand'].map(eff_dict)
                
                df_melted = df_melted.dropna(subset=['Price'])
                all_data.append(df_melted)
                success_for_this_day = True
                print(f"  ✅ ดึงข้อมูลวันที่ {date_str} สำเร็จเรียบร้อย")
                break
                
            except Exception as e:
                print(f"  -> รอบที่ {attempt} พบข้อผิดพลาด: {e}")
                time.sleep(2)
                
        if not success_for_this_day:
            print(f"  ❌ ข้ามวันที่ {date_str} เนื่องจากไม่สามารถดึงข้อมูลได้ครบถ้วน")

# ------------------ 3. รวมและบันทึกลง Parquet ------------------
if all_data:
    new_df = pd.concat(all_data, ignore_index=True)
    new_df['Date'] = pd.to_datetime(new_df['Date'], format='%d/%m/%Y')

    if existing_df is not None and not existing_df.empty:
        master_df = pd.concat([existing_df, new_df], ignore_index=True)
        master_df.drop_duplicates(subset=['Date', 'Brand', 'Product'], keep='last', inplace=True)
    else:
        master_df = new_df

    master_df.sort_values(by=['Brand', 'Date', 'Product'], inplace=True)
    master_df.to_parquet(output_file, engine='pyarrow', index=False)
    print(f"\n🎉 บันทึกไฟล์สำเร็จ: {output_file} (รวม {len(master_df):,} แถว)")
else:
    print("\nไม่มีข้อมูลใหม่ถูกดึงมาเพิ่มเติม")
