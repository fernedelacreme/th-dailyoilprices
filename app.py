import glob
from io import BytesIO
import os
from generate_report import build_monthly_report
import pandas as pd
import requests
import streamlit as st

# ------------------ 1. ตั้งค่าหน้าเว็บ ------------------
st.set_page_config(
    page_title="ระบบสรุปรายงานราคาน้ำมันขายปลีก", page_icon="⛽", layout="wide"
)

# ------------------ ปรับแต่งสไตล์และฟอนต์ผ่าน Google Fonts (ป้องกัน Icon พัง) ------------------
custom_font_css = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Kanit:wght@400;500;600&family=Prompt:wght@300;400;500;600&display=swap');

    html, body, p, span:not([data-testid="stIconMaterial"]):not([class*="material-symbols"]), 
    div:not([data-testid="stIconMaterial"]):not([class*="material-symbols"]), 
    label, input, select {
        font-family: 'Prompt', sans-serif !important;
    }

    h1, h2, h3, h4, h5, h6 {
        font-family: 'Kanit', sans-serif !important;
        font-weight: 500 !important;
    }

    .stButton > button, .stDownloadButton > button {
        font-family: 'Prompt', sans-serif !important;
        font-weight: 500;
        border-radius: 8px;
    }

    [data-testid="stIconMaterial"], 
    [class*="material-symbols"], 
    i[class*="icon"], 
    span[class*="material"] {
        font-family: 'Material Symbols Rounded', 'Material Symbols Outlined', 'Material Icons' !important;
    }
</style>
"""
st.markdown(custom_font_css, unsafe_allow_html=True)

st.title("⛽ ระบบรายงานราคาน้ำมันขายปลีกอัตโนมัติ")
st.markdown("""
    ระบบแปลงข้อมูลราคาน้ำมันขายปลีกจาก EPPO จัดทำรายงาน 3 ชีท:
    1. **ราคาน้ำมันขายปลีก**
    2. **ราคารวมหัวจ่าย (+0.05 บาท)**
    3. **ผลต่างราคารายวัน**
    """)
st.write("---")

data_dir = "data"


# ------------------ 2. ฟังก์ชันสั่งรัน GitHub Actions ผ่าน API ------------------
def trigger_github_workflow():
  if (
      "GH_TOKEN" not in st.secrets
      or "REPO_OWNER" not in st.secrets
      or "REPO_NAME" not in st.secrets
  ):
    return (
        False,
        "ยังไม่ได้ตั้งค่า Secrets (GH_TOKEN, REPO_OWNER, REPO_NAME) ใน"
        " Streamlit Cloud",
    )

  token = st.secrets["GH_TOKEN"]
  owner = st.secrets["REPO_OWNER"]
  repo = st.secrets["REPO_NAME"]
  workflow_file = "daily_scrape.yml"

  url = f"https://api.github.com/repos/{owner}/{repo}/actions/workflows/{workflow_file}/dispatches"
  headers = {
      "Authorization": f"Bearer {token}",
      "Accept": "application/vnd.github+json",
  }
  payload = {"ref": "main"}

  try:
    res = requests.post(url, json=payload, headers=headers, timeout=15)
    if res.status_code == 204:
      return (
          True,
          "ส่งคำสั่งสำเร็จ! ระบบกำลังดึงข้อมูลบน GitHub Actions"
          " (ใช้เวลาประมาณ 1–2 นาที)",
      )
    else:
      return (
          False,
          f"ส่งคำสั่งไม่สำเร็จ (HTTP Status: {res.status_code}): {res.text}",
      )
  except Exception as e:
    return False, f"เกิดข้อผิดพลาดในการเชื่อมต่อ: {e}"


# ------------------ 3. โหลดข้อมูลฐานข้อมูล Parquet แยกรายปี ------------------
@st.cache_data(ttl=3600)
def load_all_data():
  parquet_files = glob.glob(os.path.join(data_dir, "EPPO_Prices_*.parquet"))
  if not parquet_files:
    return pd.DataFrame()

  df_list = []
  for f in parquet_files:
    try:
      df_temp = pd.read_parquet(f)
      df_list.append(df_temp)
    except Exception as e:
      st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์ {f}: {e}")

  if not df_list:
    return pd.DataFrame()

  full_df = pd.concat(df_list, ignore_index=True)
  full_df["Date"] = pd.to_datetime(full_df["Date"])
  full_df.sort_values(by=["Brand", "Date", "Product"], inplace=True)
  return full_df


df_raw = load_all_data()

# ------------------ ตรวจสอบว่ามีข้อมูลหรือไม่ ------------------
if df_raw.empty:
  st.error(
      f"⚠️ ไม่พบไฟล์ข้อมูล Parquet ในโฟลเดอร์ `{data_dir}/`\n\n"
      "กรุณาตรวจสอบว่ามีไฟล์ `data/EPPO_Prices_YYYY.parquet` บน Repository"
      " หรือกดสั่งดึงข้อมูลจากเมนูด้านซ้าย"
  )
  st.sidebar.header("🔄 อัปเดตข้อมูล")
  if st.sidebar.button("🚀 สั่งดึงข้อมูลรอบแรก"):
    success, msg = trigger_github_workflow()
    if success:
      st.sidebar.success(msg)
    else:
      st.sidebar.error(msg)
  st.stop()

# ------------------ 4. ส่วน Sidebar (สถานะระบบ & ดาวน์โหลด Raw & ปุ่มสั่งอัปเดต) ------------------
min_date = df_raw["Date"].min()
max_date = df_raw["Date"].max()

st.sidebar.header("📌 ข้อมูลสถานะระบบ")
st.sidebar.info(
    f"**ช่วงข้อมูลที่มีในระบบ:**\n\n"
    f"- เริ่มต้น: {min_date.strftime('%d/%m/%Y')}\n"
    f"- ล่าสุด: {max_date.strftime('%d/%m/%Y')}\n\n"
    f"**จำนวนแถวทั้งหมด:** {len(df_raw):,} แถว\n"
    f"**การจัดเก็บ:** Parquet (แยกรายปี)"
)

# ส่วนดาวน์โหลด Raw Data (แปลงเป็น Excel ให้โหลดทางหน้าเว็บ)
st.sidebar.write("---")
st.sidebar.subheader("📥 ดาวน์โหลดฐานข้อมูลดิบ (Raw Data)")
raw_dl_option = st.sidebar.radio(
    "เลือกรูปแบบข้อมูลดิบที่ต้องการดาวน์โหลด:",
    options=["เฉพาะปีล่าสุด", "รวมทั้งหมดทุกปี"],
)

if raw_dl_option == "เฉพาะปีล่าสุด":
  latest_year = max_date.year
  export_raw_df = df_raw[df_raw["Date"].dt.year == latest_year]
  raw_filename = f"EPPO_Raw_Prices_{latest_year}.xlsx"
else:
  export_raw_df = df_raw
  raw_filename = "EPPO_Raw_Prices_All_Years.xlsx"


@st.cache_data
# ปรับฟังก์ชันแปลงไฟล์ ไม่ต้องใช้แคชถาวร หรือระบุ ttl สั้นๆ เพื่อให้ได้ไฟล์สดใหม่เสมอ
def convert_df_to_excel(df_to_export):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl', datetime_format='DD/MM/YYYY') as writer:
        df_to_export.to_excel(writer, index=False)
    return output.getvalue()

# สั่งแปลงเป็น Excel โดยตรงตาม DataFrame ก้อนล่าสุด
excel_data = convert_df_to_excel(export_raw_df)

st.sidebar.download_button(
    label=f"💾 ดาวน์โหลด {raw_filename}",
    data=excel_data,
    file_name=raw_filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="btn_download_raw"
)

# ส่วนปุ่ม Trigger สั่งดึงข้อมูล
st.sidebar.write("---")
st.sidebar.subheader("🔄 ตรวจสอบข้อมูลล่าสุด")

if st.sidebar.button("🧹 ล้างแคชและรีเฟรชหน้าเว็บ"):
  st.cache_data.clear()
  st.rerun()

import os
from datetime import datetime, timedelta, timezone

parquet_path = "data/EPPO_Prices_2026.parquet"
if os.path.exists(parquet_path):
    # ดึงเวลา modified ของไฟล์ แล้วแปลงเป็นเวลาไทย (UTC+7)
    mod_timestamp = os.path.getmtime(parquet_path)
    tz_th = timezone(timedelta(hours=7))
    mod_time = datetime.fromtimestamp(mod_timestamp, tz=timezone.utc).astimezone(tz_th)
    
    st.sidebar.caption(f"🕒 ข้อมูลล่าสุดในไฟล์: {mod_time.strftime('%d/%m/%Y %H:%M น.')}")

# ------------------ 5. ฟอร์มเลือกเดือนและปีสำหรับออกรายงาน ------------------
month_names_th = {
    1: "มกราคม",
    2: "กุมภาพันธ์",
    3: "มีนาคม",
    4: "เมษายน",
    5: "พฤษภาคม",
    6: "มิถุนายน",
    7: "กรกฎาคม",
    8: "สิงหาคม",
    9: "กันยายน",
    10: "ตุลาคม",
    11: "พฤศจิกายน",
    12: "ธันวาคม",
}

st.subheader("🗓️ เลือกงวดเดือนและปีที่ต้องการออกรายงาน")

col1, col2 = st.columns(2)

available_years = sorted(df_raw["Date"].dt.year.unique(), reverse=True)

with col1:
  selected_year = st.selectbox(
      "เลือกปี (พ.ศ. / ค.ศ.)",
      options=available_years,
      format_func=lambda y: f"พ.ศ. {y + 543} (ค.ศ. {y})",
  )

available_months = sorted(
    df_raw[df_raw["Date"].dt.year == selected_year]["Date"].dt.month.unique()
)
if not available_months:
  available_months = list(range(1, 13))

with col2:
  selected_month = st.selectbox(
      "เลือกเดือน",
      options=available_months,
      format_func=lambda m: f"{month_names_th.get(m, m)} (เดือน {m})",
  )

# ------------------ 6. ปุ่มสร้างและดาวน์โหลดไฟล์รายงาน (ใช้ Session State ป้องกันปุ่มหาย) ------------------
if "generated_report" not in st.session_state:
  st.session_state.generated_report = None

st.write("")
if st.button("📊 ประมวลผลและสร้างไฟล์รายงาน Excel", type="primary"):
  filtered_df = df_raw[
      (df_raw["Date"].dt.year == selected_year)
      & (df_raw["Date"].dt.month == selected_month)
  ]

  if filtered_df.empty:
    st.warning(
        f"⚠️ ไม่พบข้อมูลราคาของเดือน {month_names_th[selected_month]}"
        f" {selected_year + 543} ในฐานข้อมูล"
    )
    st.session_state.generated_report = None
  else:
    with st.spinner("กำลังคำนวณและจัดรูปแบบตาราง Excel ทั้ง 3 ชีท..."):
      output_filename = (
          f"รายงานราคาน้ำมัน_{selected_year}_{selected_month:02d}.xlsx"
      )

      try:
        # สั่งสร้างรีพอร์ต 3 ชีทโดยใช้ฟังก์ชันใน generate_report.py
        build_monthly_report(
            df_raw, selected_year, selected_month, output_filename
        )

        with open(output_filename, "rb") as f:
          file_bytes = f.read()

        # บันทึกลง Session State เพื่อคงปุ่มดาวน์โหลดไว้
        st.session_state.generated_report = {
            "bytes": file_bytes,
            "filename": output_filename,
            "month_name": month_names_th[selected_month],
            "year_th": selected_year + 543,
            "preview_df": filtered_df.sort_values(
                by=["Date", "Product", "Brand"]
            ),
        }

      except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาดระหว่างการสร้างรายงาน: {e}")
        st.session_state.generated_report = None

# แสดงผลปุ่มดาวน์โหลดค้างไว้เสมอ
if st.session_state.generated_report is not None:
  rep = st.session_state.generated_report
  st.success(
      f"✅ ไฟล์รายงานประจำเดือน {rep['month_name']} {rep['year_th']}"
      " พร้อมดาวน์โหลดแล้ว"
  )

  st.download_button(
      label="📥 คลิกที่นี่เพื่อดาวน์โหลดไฟล์ Excel (.xlsx)",
      data=rep["bytes"],
      file_name=rep["filename"],
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      key="btn_download_report",
  )

  with st.expander("👀 ดูตัวอย่างข้อมูลดิบเฉพาะงวดเดือนที่เลือก"):
    st.dataframe(rep["preview_df"], use_container_width=True)
