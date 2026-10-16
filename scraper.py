import os
from datetime import datetime, timedelta, timezone
from io import BytesIO
import pandas as pd
import requests
from bs4 import BeautifulSoup

# ------------------ ตั้งค่าโฟลเดอร์และตัวแปร ------------------
data_dir = 'data'
os.makedirs(data_dir, exist_ok=True)

target_url = "https://www2.eppo.go.th/epporop/entryreport/ropbydaypublic.aspx"

tz_th = timezone(timedelta(hours=7))
today = datetime.now(tz_th).date()
current_year = today.year
output_file = os.path.join(data_dir, f'EPPO_Prices_{current_year}.parquet')

# ดึง ScraperAPI Key จาก GitHub Secret
scraper_api_key = os.environ.get("SCRAPER_API_KEY", "").strip()

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

# ------------------ 2. ดึงข้อมูลราคาน้ำมันจาก EPPO ------------------
with requests.Session() as session:
    # หากมี ScraperAPI Key ให้ตั้งค่าเป็น Standard Proxy (แก้ปัญหา Error 500)
    if scraper_api_key:
        proxies = {
            "http": f"http://scraperapi:{scraper_api_key}@proxy-server.scraperapi.com:8001",
            "https": f"http://scraperapi:{scraper_api_key}@proxy-server.scraperapi.com:8001"
        }
        session.proxies.update(proxies)
        # ปิดการตรวจ Verify SSL เมื่อวิ่งผ่าน Proxy
        session.verify = False
    else:
        session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36'
        })

    for date_str in date_list:
        print(f"กำลังดึงและแปลงข้อมูลวันที่: {date_str}")
        try:
            # 1. ยิง GET ตรงๆ ไปที่ URL ปลายทาง (Session จะพาผ่าน ScraperAPI Proxy อัตโนมัติ)
            res_get = session.get(target_url, timeout=90)
            
            if res_get.status_code != 200:
                print(f"  -> ยิง GET ไม่สำเร็จ (Status Code: {res_get.status_code})")
                continue

            soup = BeautifulSoup(res_get.text, 'html.parser')
            viewstate = soup.find('input', {'name': '__VIEWSTATE'})
            viewstate_gen = soup.find('input', {'name': '__VIEWSTATEGENERATOR'})
            event_val = soup.find('input', {'name': '__EVENTVALIDATION'})

            if not viewstate or not event_val:
                print(f"  -> ไม่พบ ViewState/EventValidation สำหรับวันที่ {date_str}")
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
            res_post = session.post(target_url, data=payload, timeout=90)

            if len(res_post.content) < 500:
                print(f"  -> ขนาดไฟล์ Excel เล็กผิดปกติ อาจยังไม่มีข้อมูลวันที่ {date_str}")
                continue

            # 3. แปลง Excel เป็น DataFrame
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
            print(f"  ✅ ดึงข้อมูลวันที่ {date_str} สำเร็จ ({len(df_melted)} รายการ)")

        except Exception as e:
            print(f"  -> ข้ามวันที่ {date_str} เนื่องจากข้อผิดพลาด: {e}")

# ------------------ 3. รวมและบันทึกลงไฟล์ Parquet ------------------
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
    print(f"\n🎉 อัปเดตและบันทึกข้อมูลเรียบร้อยแล้ว: {output_file} (รวมทั้งหมด {len(master_df):,} แถว)")
else:
    print("\nไม่มีข้อมูลใหม่ถูกดึงมาเพิ่มเติม")
