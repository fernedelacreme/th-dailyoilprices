import os
from io import BytesIO
import requests
import pandas as pd
import streamlit as st
from generate_report import build_monthly_report

# ------------------ 1. ตั้งค่าหน้าเว็บ ------------------
st.set_page_config(
    page_title="ระบบสรุปรายงานราคาน้ำมันขายปลีก",
    page_icon="⛽",
    layout="wide"
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
st.markdown(
    """
    ระบบแปลงข้อมูลราคาน้ำมันขายปลีกจาก EPPO จัดทำรายงาน 3 ชีท:
    1. **ราคาน้ำมันขายปลีก**
    2. **ราคารวมหัวจ่าย (+0.05 บาท)**
    3. **ผลต่างราคารายวัน**
    """
)
st.write("---")

raw_file = "EPPO_Historical_Prices_Merged.xlsx"

# ------------------ 2. ฟังก์ชันสั่งรัน GitHub Actions ผ่าน API ------------------
def trigger_github_workflow():
    if "GH_TOKEN" not in st.secrets or "REPO_OWNER" not in st.secrets or "REPO_NAME" not in st.secrets:
        return False, "ยังไม่ได้ตั้งค่า Secrets (GH_TOKEN, REPO_OWNER, REPO_NAME) ใน Streamlit Cloud"
    
    token = st.secrets["GH_TOKEN"]
    owner = st.secrets["REPO_OWNER"]
    repo = st.secrets["REPO_NAME"]
    workflow_file = "daily_scrape.yml"

    url = f"https://api.github.com/repos/{owner}/{repo}/actions/workflows/{workflow_file}/dispatches"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json"
    }
    payload = {
        "ref": "main"
    }

    try:
        res = requests.post(url, json=payload, headers=headers, timeout=15)
        if res.status_code == 204:
            return True, "ส่งคำสั่งสำเร็จ! ระบบกำลังดึงข้อมูลบน GitHub Actions (ใช้เวลาประมาณ 1–2 นาที)"
        else:
            return False, f"ส่งคำสั่งไม่สำเร็จ (HTTP Status: {res.status_code}): {res.text}"
    except Exception as e:
        return False, f"เกิดข้อผิดพลาดในการเชื่อมต่อ: {e}"

# ------------------ 3. ตรวจสอบไฟล์ฐานข้อมูล ------------------
if not os.path.exists(raw_file):
    st.error(
        f"⚠️ ไม่พบไฟล์ฐานข้อมูล `{raw_file}` ในระบบ\n\n"
        "กรุณารอระบบ GitHub Actions ทำการ Scrape ข้อมูลในรอบแรก หรือกดสั่งรันข้อมูลจากเมนูด้านซ้าย"
    )
    st.sidebar.header("🔄 อัปเดตข้อมูล")
    if st.sidebar.button("🚀 สั่งดึงข้อมูลรอบแรก"):
        success, msg = trigger_github_workflow()
        if success:
            st.sidebar.success(msg)
        else:
            st.sidebar.error(msg)
    st.stop()

# ------------------ 4. โหลดข้อมูลฐานข้อมูล ------------------
@st.cache_data(ttl=3600)
def load_data(filepath):
    df = pd.read_excel(filepath)
    df['Date'] = pd.to_datetime(df['Date'])
    return df

df_raw = load_data(raw_file)

if df_raw.empty:
    st.warning("⚠️ ไฟล์ฐานข้อมูลว่างเปล่า ยังไม่มีรายการราคาน้ำมัน")
    st.stop()

# ------------------ 5. ส่วน Sidebar (สถานะระบบ & ดาวน์โหลด Raw & ปุ่มสั่งอัปเดต) ------------------
min_date = df_raw['Date'].min()
max_date = df_raw['Date'].max()

st.sidebar.header("📌 ข้อมูลสถานะระบบ")
st.sidebar.info(
    f"**ช่วงข้อมูลที่มีในระบบ:**\n\n"
    f"- เริ่มต้น: {min_date.strftime('%d/%m/%Y')}\n"
    f"- ล่าสุด: {max_date.strftime('%d/%m/%Y')}\n\n"
    f"**จำนวนแถวทั้งหมด:** {len(df_raw):,} แถว"
)

# ปุ่มดาวน์โหลด Raw Data
st.sidebar.write("---")
st.sidebar.subheader("📥 ดาวน์โหลดฐานข้อมูลดิบ (Raw Data)")
st.sidebar.caption("ไฟล์ Excel รวมราคาน้ำมันทั้งหมดที่เก็บสะสมตั้งแต่เริ่มต้นจนถึงปัจจุบัน")

with open(raw_file, "rb") as rf:
    raw_bytes = rf.read()

st.sidebar.download_button(
    label="💾 ดาวน์โหลด EPPO Raw Data (.xlsx)",
    data=raw_bytes,
    file_name="EPPO_Historical_Prices_Merged.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="btn_download_raw"
)

# ส่วน Trigger ดึงข้อมูลใหม่
st.sidebar.write("---")
st.sidebar.subheader("🔄 สั่งดึงข้อมูลล่าสุด")
st.sidebar.caption("กดปุ่มนี้เพื่อส่งสัญญาณให้ GitHub Actions ดึงข้อมูลวันที่ขาดอยู่จนถึงปัจจุบันทันที")

if st.sidebar.button("🚀 อัปเดตข้อมูลเดี๋ยวนี้", type="secondary"):
    success, msg = trigger_github_workflow()
    if success:
        st.sidebar.success(msg)
        st.sidebar.info("💡 รอประมาณ 1–2 นาทีให้บอททำงานเสร็จ แล้วกดปุ่ม **'ล้างแคชและรีเฟรช'** ด้านล่าง")
    else:
        st.sidebar.error(msg)

if st.sidebar.button("🧹 ล้างแคชและรีเฟรชหน้าเว็บ"):
    st.cache_data.clear()
    st.rerun()

# ------------------ 6. ฟอร์มเลือกเดือนและปีสำหรับออกรายงาน ------------------
month_names_th = {
    1: "มกราคม", 2: "กุมภาพันธ์", 3: "มีนาคม", 4: "เมษายน",
    5: "พฤษภาคม", 6: "มิถุนายน", 7: "กรกฎาคม", 8: "สิงหาคม",
    9: "กันยายน", 10: "ตุลาคม", 11: "พฤศจิกายน", 12: "ธันวาคม"
}

st.subheader("🗓️ เลือกงวดเดือนและปีที่ต้องการออกรายงาน")

col1, col2 = st.columns(2)

available_years = sorted(df_raw['Date'].dt.year.unique(), reverse=True)

with col1:
    selected_year = st.selectbox(
        "เลือกปี (พ.ศ. / ค.ศ.)",
        options=available_years,
        format_func=lambda y: f"พ.ศ. {y + 543} (ค.ศ. {y})"
    )

available_months = sorted(df_raw[df_raw['Date'].dt.year == selected_year]['Date'].dt.month.unique())
if not available_months:
    available_months = list(range(1, 13))

with col2:
    selected_month = st.selectbox(
        "เลือกเดือน",
        options=available_months,
        format_func=lambda m: f"{month_names_th.get(m, m)} (เดือน {m})"
    )

# ------------------ 7. ปุ่มสร้างและดาวน์โหลดไฟล์รายงาน (ใช้ Session State) ------------------
if "generated_report" not in st.session_state:
    st.session_state.generated_report = None

st.write("")
if st.button("📊 ประมวลผลและสร้างไฟล์รายงาน Excel", type="primary"):
    filtered_df = df_raw[(df_raw['Date'].dt.year == selected_year) & (df_raw['Date'].dt.month == selected_month)]
    
    if filtered_df.empty:
        st.warning(f"⚠️ ไม่พบข้อมูลราคาของเดือน {month_names_th[selected_month]} {selected_year + 543} ในฐานข้อมูล")
        st.session_state.generated_report = None
    else:
        with st.spinner("กำลังคำนวณและจัดรูปแบบตาราง Excel ทั้ง 3 ชีท..."):
            output_filename = f"รายงานราคาน้ำมัน_{selected_year}_{selected_month:02d}.xlsx"
            
            try:
                build_monthly_report(df_raw, selected_year, selected_month, output_filename)
                
                with open(output_filename, "rb") as f:
                    file_bytes = f.read()

                # บันทึกข้อมูลลง Session State เพื่อไม่ให้ปุ่มดาวน์โหลดหายหลังกดคลิก
                st.session_state.generated_report = {
                    "bytes": file_bytes,
                    "filename": output_filename,
                    "month_name": month_names_th[selected_month],
                    "year_th": selected_year + 543,
                    "preview_df": filtered_df.sort_values(by=['Date', 'Product', 'Brand'])
                }

            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาดระหว่างการสร้างรายงาน: {e}")
                st.session_state.generated_report = None

# แสดงผลปุ่มดาวน์โหลดค้างไว้เสมอ
if st.session_state.generated_report is not None:
    rep = st.session_state.generated_report
    st.success(f"✅ ไฟล์รายงานประจำเดือน {rep['month_name']} {rep['year_th']} พร้อมดาวน์โหลดแล้ว")
    
    st.download_button(
        label="📥 คลิกที่นี่เพื่อดาวน์โหลดไฟล์ Excel (.xlsx)",
        data=rep["bytes"],
        file_name=rep["filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="btn_download_report"
    )

    with st.expander("👀 ดูตัวอย่างข้อมูลดิบเฉพาะงวดเดือนที่เลือก"):
        st.dataframe(
            rep["preview_df"],
            use_container_width=True
        )
