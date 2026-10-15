import os
from datetime import datetime, timedelta
import pandas as pd
import requests
from bs4 import BeautifulSoup
from io import BytesIO

output_file = 'EPPO_Historical_Prices_Merged.xlsx'
url = "https://www2.eppo.go.th/epporop/entryreport/ropbydaypublic.aspx"
default_start_date = '2025-01-01'

today = datetime.today().date()
existing_df = None

if os.path.exists(output_file):
    existing_df = pd.read_excel(output_file)
    existing_df['Date'] = pd.to_datetime(existing_df['Date'])
    if not existing_df.empty and existing_df['Date'].notna().any():
        last_date = existing_df['Date'].max().date()
        start_date = last_date + timedelta(days=1)
    else:
        start_date = pd.to_datetime(default_start_date).date()
else:
    start_date = pd.to_datetime(default_start_date).date()

if start_date > today:
    print(f"ข้อมูลในไฟล์เป็นปัจจุบันแล้ว (ถึง {today}) ไม่ต้องดึงเพิ่ม")
    exit(0)

print(f"กำลังดึงข้อมูลช่วง {start_date} ถึง {today}")
date_list = pd.date_range(start=start_date, end=today).strftime('%d/%m/%Y')
all_data = []

with requests.Session() as session:
    for date_str in date_list:
        try:
            res_get = session.get(url, timeout=30)
            soup = BeautifulSoup(res_get.text, 'html.parser')
            
            payload = {
                '__EVENTTARGET': '',
                '__EVENTARGUMENT': '',
                '__VIEWSTATE': soup.find('input', {'name': '__VIEWSTATE'})['value'],
                '__VIEWSTATEGENERATOR': soup.find('input', {'name': '__VIEWSTATEGENERATOR'})['value'],
                '__EVENTVALIDATION': soup.find('input', {'name': '__EVENTVALIDATION'})['value'],
                'TbxToDate': date_str,
                'BtnGenerate.x': '50',
                'BtnGenerate.y': '10'
            }
            
            res_post = session.post(url, data=payload, timeout=30)
            df_raw = pd.read_excel(BytesIO(res_post.content), header=None)
            
            brands = df_raw.iloc[11, 1:11].fillna('').astype(str).tolist()
            products = df_raw.iloc[13:25, 0].fillna('').astype(str).tolist()
            
            prices = df_raw.iloc[13:25, 1:11].copy()
            prices.columns = brands
            prices['Product'] = products
            prices['Date'] = date_str
            
            df_melted = prices.melt(id_vars=['Date', 'Product'], value_vars=brands, var_name='Brand', value_name='Price')
            
            eff_dates = df_raw.iloc[25, 1:11].fillna('').astype(str).tolist()
            df_melted['Effective_Date'] = df_melted['Brand'].map(dict(zip(brands, eff_dates)))
            df_melted = df_melted.dropna(subset=['Price'])
            all_data.append(df_melted)
        except Exception as e:
            print(f"ข้าม {date_str}: {e}")

if all_data:
    new_df = pd.concat(all_data, ignore_index=True)
    new_df['Date'] = pd.to_datetime(new_df['Date'], format='%d/%m/%Y')
    
    if existing_df is not None and not existing_df.empty:
        master_df = pd.concat([existing_df, new_df], ignore_index=True)
        master_df.drop_duplicates(subset=['Date', 'Brand', 'Product'], keep='last', inplace=True)
    else:
        master_df = new_df
        
    master_df.sort_values(by=['Brand', 'Date', 'Product'], inplace=True)
    with pd.ExcelWriter(output_file, engine='openpyxl', datetime_format='DD/MM/YYYY') as writer:
        master_df.to_excel(writer, index=False)
    print("อัปเดตไฟล์ Excel สำเร็จ")