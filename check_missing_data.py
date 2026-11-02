import os
from datetime import datetime, timedelta, timezone
import pandas as pd

# 1. ตั้งค่าเวลาปัจจุบัน
tz_th = timezone(timedelta(hours=7))
today = datetime.now(tz_th).date()
current_year = today.year

# รายชื่อปีที่ต้องการเช็ก
years_to_check = [2025, current_year]

print("📊 รายงานการตรวจสอบความสมบูรณ์ของไฟล์ราคาน้ำมัน EPPO")
print("=" * 60)

for year in years_to_check:
    parquet_file = os.path.join('data', f'EPPO_Prices_{year}.parquet')
    start_of_year = datetime(year, 1, 1).date()
    end_of_check = datetime(year, 12, 31).date() if year < current_year else today

    if not os.path.exists(parquet_file):
        print(f"\n❌ [ปี {year}] ไม่พบไฟล์ {parquet_file}")
        continue

    # อ่านข้อมูล Parquet
    df = pd.read_parquet(parquet_file)
    df['Date'] = pd.to_datetime(df['Date'])

    # คำนวณช่วงวันที่ต้องมี
    expected_dates = pd.date_range(start=start_of_year, end=end_of_check).date
    existing_dates = df['Date'].dt.date.unique()
    missing_dates = [d for d in expected_dates if d not in existing_dates]

    # แสดงผลลัพธ์ประจำปี
    print(f"\n🗓️  ปี {year} ({start_of_year.strftime('%d/%m/%Y')} ถึง {end_of_check.strftime('%d/%m/%Y')})")
    print(f"    - จำนวนวันที่ควรมีทั้งหมด: {len(expected_dates)} วัน")
    print(f"    - มีข้อมูลอยู่ในไฟล์แล้ว: {len(existing_dates)} วัน")

    if missing_dates:
        print(f"    ⚠️  พบวันที่ขาดหายไป {len(missing_dates)} วัน ได้แก่:")
        for d in missing_dates:
            print(f"        • {d.strftime('%d/%m/%Y')}")
    else:
        print(f"    ✅ สมบูรณ์ 100%! ไม่พบวันที่ขาดหายไป")

print("\n" + "=" * 60)