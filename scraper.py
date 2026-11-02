import os
import urllib3
from datetime import datetime, timedelta, timezone
from io import BytesIO
import pandas as pd
import requests
from bs4 import BeautifulSoup
import sys
import io

# บังคับให้ stdout รองรับ UTF-8 และละเว้นหรือแทนที่ตัวอักษรที่แปลงไม่ได้
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# ปิดคำเตือน SSL เผื่อใบรับรองเว็บราชการมีปัญหา
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ------------------ ตั้งค่าโฟลเดอร์และตัวแปร ------------------
data_dir = 'data'
os.makedirs(data_dir, exist_ok=True)

target_url = "https://www2.eppo.go.th/epporop/entryreport/ropbydaypublic.aspx"

tz_th = timezone(timedelta(hours=7))
today = datetime.now(tz_th).date()
current_year = today.year
output_file = os.path.join(data_dir, f'EPPO_Prices_{current_year}.parquet')

# ------------------ 1. ตรวจสอบข้อมูลเดิมและหาวันที่ขาดหายไป (Backfill Logic) ------------------
existing_df = None
start_of_year = datetime(current_year, 1, 1).date()

if os.path.exists(output_file):
    print(f"กำลังอ่านข้อมูลเดิมจาก '{output_file}'...")
    try:
        existing_df = pd.read_parquet(output_file)
        existing_df['Date'] = pd.to_datetime(existing_df['Date'])

        # สร้างช่วงวันที่ควรจะมีทั้งหมดตั้งแต่ต้นปีถึงวันนี้
        expected_dates = pd.date_range(start=start_of_year, end=today).date

        # ดึงรายการวันที่มีอยู่แล้วในไฟล์ Parquet
        existing_dates = existing_df['Date'].dt.date.unique()

        # หาวันที่ยังขาดหายไปในอดีต (Missing Dates)
        missing_dates = [d for d in expected_dates if d not in existing_dates]

        if not missing_dates:
            print(f"✅ ข้อมูลปี {current_year} สมบูรณ์ครบถ้วนทุกวันแล้ว (ถึง {today.strftime('%d/%m/%Y')})")
            sys.exit(0)

        print(f"พบวันที่ข้อมูลขาดหายไปทั้งหมด {len(missing_dates)} วัน")
        date_list = [d.strftime('%d/%m/%Y') for d in missing_dates]
    except Exception as e:
        print(f"⚠️ ไม่สามารถอ่านไฟล์เดิมได้ ({e}) จะเริ่มดึงใหม่ทั้งหมดตั้งแต่ต้นปี...")
        date_list = pd.date_range(start=start_of_year, end=today).strftime('%d/%m/%Y')
else:
    print(f"ไม่พบไฟล์ปี {current_year} จะเริ่มดึงตั้งแต่วันที่ 01/01/{current_year}...")
    date_list = pd.date_range(start=start_of_year, end=today).strftime('%d/%m/%Y')

print(f"ช่วงวันที่ต้องดึงข้อมูลมาเติม: {len(date_list)} วัน\n")
all_data = []

# ------------------ 2. ดึงข้อมูลราคาน้ำมันจาก EPPO โดยตรง ------------------
with requests.Session() as session:
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36'
    })
    session.verify = False

    # 2.1 ยิง GET ครั้งเดียวก่อนเข้าลูปเพื่อเก็บ ViewState
    try:
        res_get = session.get(target_url, timeout=60)
        if res_get.status_code != 200:
            print(f"❌ เชื่อมต่อหน้าเว็บ EPPO ไม่สำเร็จ (Status: {res_get.status_code})")
            sys.exit(1)

        soup = BeautifulSoup(res_get.text, 'html.parser')
        viewstate = soup.find('input', {'name': '__VIEWSTATE'})
        viewstate_gen = soup.find('input', {'name': '__VIEWSTATEGENERATOR'})
        event_val = soup.find('input', {'name': '__EVENTVALIDATION'})

        if not viewstate or not event_val:
            print("❌ ไม่พบ ViewState ในหน้าเว็บ")
            sys.exit(1)
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดในการเชื่อมต่อเริ่มต้น: {e}")
        sys.exit(1)

   # 2.2 วนลูปยิง POST เฉพาะวันที่ขาดหายไป
    for date_str in date_list:
        print(f"กำลังดึงและแปลงข้อมูลวันที่: {date_str}")
        
        # ลองยิงซ้ำสูงสุด 3 ครั้งเผื่อเน็ตกระตุก/IncompleteRead
        max_retries = 3
        for attempt in range(1, max_retries + 1):
            try:
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

                res_post = session.post(target_url, data=payload, timeout=60)

                if len(res_post.content) < 500 or b'<html' in res_post.content[:200].lower():
                    print(f"  -> อาจยังไม่มีข้อมูลสำหรับวันที่ {date_str} (ข้าม)")
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
                print(f"  ✅ ดึงข้อมูลวันที่ {date_str} สำเร็จ ({len(df_melted)} รายการ)")
                break  # สำเร็จแล้ว หลุดออกจากลูป retry

            except Exception as e:
                if attempt < max_retries:
                    print(f"  ⚠️ ดึงวันที่ {date_str} ไม่สำเร็จ ({e}) กำลังลองใหม่รอบที่ {attempt + 1}...")
                else:
                    print(f"  ❌ ข้ามวันที่ {date_str} เนื่องจากลองครบ {max_retries} ครั้งแล้วยังล้มเหลว")

# ------------------ 3. รวมและบันทึกลงไฟล์ Parquet ------------------
if all_data:
    new_df = pd.concat(all_data, ignore_index=True)
    new_df['Date'] = pd.to_datetime(new_df['Date'], format='%d/%m/%Y')

    if existing_df is not None and not existing_df.empty:
        master_df = pd.concat([existing_df, new_df], ignore_index=True)
        # ลบข้อมูลซ้ำ โดยยึดเอาข้อมูลล่าสุดที่เพิ่งดึงมา
        master_df.drop_duplicates(subset=['Date', 'Brand', 'Product'], keep='last', inplace=True)
    else:
        master_df = new_df

    # เรียงลำดับตามแบรนด์ วันที่ และชื่อน้ำมัน
    master_df.sort_values(by=['Brand', 'Date', 'Product'], inplace=True)
    master_df.to_parquet(output_file, engine='pyarrow', index=False)
    print(f"\n🎉 อัปเดตและบันทึกข้อมูลเรียบร้อยแล้ว: {output_file} (รวมทั้งหมด {len(master_df):,} แถว)")
else:
    print("\nไม่มีข้อมูลใหม่ถูกดึงมาเพิ่มเติม")