import os
import time
import random
import urllib3
from datetime import datetime, timedelta, timezone
from io import BytesIO
import pandas as pd
import requests
from bs4 import BeautifulSoup

# ปิดคำเตือน InsecureRequestWarning เวลาตั้งค่า verify=False กับ Proxy
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

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


# ------------------ ฟังก์ชันช่วยดึง ViewState ตั้งต้น ------------------
def get_initial_tokens(session, url):
    """ยิง GET เพื่อดึงค่า __VIEWSTATE และ __EVENTVALIDATION จากหน้าเว็บ"""
    res = session.get(url, timeout=90)
    if res.status_code != 200:
        # ปริ้นข้อความจาก ScraperAPI เพื่อดูสาเหตุที่แท้จริงของ Error 500
        error_msg = res.text[:300].strip() if res.text else "ไม่มีข้อความตอบกลับ"
        raise Exception(f"Status Code: {res.status_code} | Detail: {error_msg}")
    
    soup = BeautifulSoup(res.text, 'html.parser')
    vs = soup.find('input', {'name': '__VIEWSTATE'})
    vs_gen = soup.find('input', {'name': '__VIEWSTATEGENERATOR'})
    ev = soup.find('input', {'name': '__EVENTVALIDATION'})
    
    if not vs or not ev:
        raise Exception("ได้หน้าเว็บมาแต่ไม่พบ ViewState (อาจติดหน้า Captcha/Firewall)")
        
    return {
        '__VIEWSTATE': vs['value'],
        '__VIEWSTATEGENERATOR': vs_gen['value'] if vs_gen else '',
        '__EVENTVALIDATION': ev['value']
    }


# ------------------ 2. ดึงข้อมูลราคาน้ำมันจาก EPPO ------------------
with requests.Session() as session:
    # ตั้งค่า Header พื้นฐานให้เหมือนเบราว์เซอร์ทั่วไป
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'th-TH,th;q=0.9,en-US;q=0.8,en;q=0.7'
    })
    session.verify = False

    tokens = None

    # 2.1 ยิง GET แค่ครั้งเดียวก่อนเข้าลูป (พร้อมระบบลองสลับโหมด Proxy อัตโนมัติ)
    if scraper_api_key:
        proxy_modes = [
            f"scraperapi.country_code=th.session_number={random.randint(10000, 999999)}",
            f"scraperapi.country_code=th.premium=true.session_number={random.randint(10000, 999999)}",
            f"scraperapi.session_number={random.randint(10000, 999999)}"
        ]

        for attempt, mode in enumerate(proxy_modes, 1):
            proxy_url = f"http://{mode}:{scraper_api_key}@proxy-server.scraperapi.com:8001"
            session.proxies.update({"http": proxy_url, "https": proxy_url})
            
            print(f"ครั้งที่ {attempt}: กำลังลองเชื่อมต่อด้วยโหมด [{mode}]...")
            try:
                tokens = get_initial_tokens(session, target_url)
                print("✅ รับค่า ViewState สำเร็จ พร้อมดึงข้อมูลรายวัน!\n")
                break
            except Exception as e:
                print(f"  ⚠️ ล้มเหลว -> {e}")
                time.sleep(3)
    else:
        try:
            print("กำลังเชื่อมต่อโดยตรง (ไม่ใช้ Proxy)...")
            tokens = get_initial_tokens(session, target_url)
            print("✅ รับค่า ViewState สำเร็จ พร้อมดึงข้อมูลรายวัน!\n")
        except Exception as e:
            print(f"  ⚠️ ล้มเหลว -> {e}")

    if not tokens:
        print("❌ ไม่สามารถเชื่อมต่อหน้าเว็บ EPPO ได้หลังจากลองทุกโหมดแล้ว")
        exit(1)

    # 2.2 วนลูปยิงเฉพาะ POST รายวันโดยใช้ค่า ViewState เดิม
    for date_str in date_list:
        print(f"กำลังดึงและแปลงข้อมูลวันที่: {date_str}")
        try:
            payload = {
                '__EVENTTARGET': '',
                '__EVENTARGUMENT': '',
                '__VIEWSTATE': tokens['__VIEWSTATE'],
                '__VIEWSTATEGENERATOR': tokens['__VIEWSTATEGENERATOR'],
                '__EVENTVALIDATION': tokens['__EVENTVALIDATION'],
                'TbxToDate': date_str,
                'BtnGenerate.x': '50',
                'BtnGenerate.y': '10'
            }

            # ยิง POST เพื่อขอไฟล์ Excel
            res_post = session.post(target_url, data=payload, timeout=90)

            # ตรวจสอบว่าผลลัพธ์ที่ได้เป็นหน้า HTML Error หรือไฟล์เล็กผิดปกติหรือไม่
            content_type = res_post.headers.get('Content-Type', '').lower()
            if 'text/html' in content_type or len(res_post.content) < 500:
                print(f"  -> ไม่ได้รับไฟล์ Excel (อาจยังไม่มีข้อมูลวันที่ {date_str} หรือ Session หมดอายุ) ลองรีเฟรช Token...")
                tokens = get_initial_tokens(session, target_url)
                payload.update(tokens)
                res_post = session.post(target_url, data=payload, timeout=90)
                
                if 'text/html' in res_post.headers.get('Content-Type', '').lower() or len(res_post.content) < 500:
                    print(f"  -> ยืนยันว่ายังไม่มีข้อมูลสำหรับวันที่ {date_str} (ข้าม)")
                    continue

            # แปลง Excel เป็น DataFrame
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
