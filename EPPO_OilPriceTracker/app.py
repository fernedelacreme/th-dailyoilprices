import os
from io import BytesIO
import pandas as pd
import streamlit as st
from generate_report import build_monthly_report

# ------------------ ตั้งค่าหน้าเว็บ ------------------
st.set_page_config(
    page_title="ระบบสรุปรายงานราคาน้ำมันขายปลีก",
    page_icon="⛽",
    layout="wide"
)

# สไตล์ส่วนหัว
st.title("⛽ ระบบรายงานราคาน้ำมันขายปลีกอัตโนมัติ")
st.markdown(
    """
    ระบบแปลงข้อมูลราคาน้ำมันดิบจาก EPPO จัดทำรายงาน 3 ชีท:
    1. **ราคาน้ำมันขายปลีก** (ตามฟอร์แมตมาตรฐาน)
    2. **ราคารวมหัวจ่าย (+0.05 บาท)**
    3. **ผลต่างราคารายวัน** (Daily Difference พร้อมค่าเฉลี่ยส่วนต่าง AI)
    """
)
st.write("---")

raw_file = "EPPO_Historical_Prices_Merged.xlsx"

# ------------------ ตรวจสอบไฟล์ฐานข้อมูล ------------------
if not os.path.exists(raw_file):
    st.error(
        f"⚠️ ไม่พบไฟล์ฐานข้อมูล `{raw_file}` ในระบบ\n\n"
        "กรุณารอระบบ GitHub Actions ทำการ Scrape ข้อมูลในรอบแรก หรืออัปโหลดไฟล์เข้าไปใน Repository"
    )
    st.stop()

# โหลดข้อมูลฐานข้อมูล
@st.cache_data(ttl=3600)
def load_data(filepath):
    df = pd.read_excel(filepath)
    df['Date'] = pd.to_datetime(df['Date'])
    return df

df_raw = load_data(raw_file)

if df_raw.empty:
    st.warning("⚠️ ไฟล์ฐานข้อมูลว่างเปล่า ยังไม่มีรายการราคาน้ำมัน")
    st.stop()

# ------------------ ส่วนควบคุม (Sidebar & Form) ------------------
min_date = df_raw['Date'].min()
max_date = df_raw['Date'].max()

st.sidebar.header("📌 ข้อมูลสถานะระบบ")
st.sidebar.info(
    f"**ช่วงข้อมูลที่มีในระบบ:**\n\n"
    f"- เริ่มต้น: {min_date.strftime('%d/%m/%Y')}\n"
    f"- ล่าสุด: {max_date.strftime('%d/%m/%Y')}\n\n"
    f"**จำนวนแถวทั้งหมด:** {len(df_raw):,} แถว"
)

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

# กรองเดือนที่มีข้อมูลจริงในปีนั้นๆ หรือให้เลือกได้ 1-12
available_months = sorted(df_raw[df_raw['Date'].dt.year == selected_year]['Date'].dt.month.unique())
if not available_months:
    available_months = list(range(1, 13))

with col2:
    selected_month = st.selectbox(
        "เลือกเดือน",
        options=available_months,
        format_func=lambda m: f"{month_names_th.get(m, m)} (เดือน {m})"
    )

# ------------------ ปุ่มสร้างรายงานและประมวลผล ------------------
st.write("")
if st.button("🚀 ประมวลผลและสร้างไฟล์รายงาน Excel", type="primary"):
    # ตรวจสอบว่าเดือนที่เลือกมีข้อมูลไหม
    filtered_df = df_raw[(df_raw['Date'].dt.year == selected_year) & (df_raw['Date'].dt.month == selected_month)]
    
    if filtered_df.empty:
        st.warning(f"⚠️ ไม่พบข้อมูลราคาของเดือน {month_names_th[selected_month]} {selected_year + 543} ในฐานข้อมูล")
    else:
        with st.spinner("กำลังประมวลผลคำนวณและจัดรูปแบบตาราง..."):
            output_filename = f"รายงานราคาน้ำมัน_{selected_year}_{selected_month:02d}.xlsx"
            
            # รันฟังก์ชันจัดรูปแบบ 3 ชีท
            try:
                build_monthly_report(df_raw, selected_year, selected_month, output_filename)
                
                # อ่านไฟล์ที่สร้างขึ้นมาใส่หน่วยความจำเพื่อให้ดาวน์โหลด
                with open(output_filename, "rb") as f:
                    file_bytes = f.read()

                st.success(f"✅ สร้างไฟล์รายงานประจำเดือน {month_names_th[selected_month]} {selected_year + 543} สำเร็จ!")
                
                # ปุ่มดาวน์โหลดไฟล์
                st.download_button(
                    label="📥 คลิกที่นี่เพื่อดาวน์โหลดไฟล์ Excel (.xlsx)",
                    data=file_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # แสดงตาราง Preview คร่าวๆ ของข้อมูลเดือนนั้น
                with st.expander("👀 ดูตัวอย่างข้อมูลดิบเฉพาะเดือนนี้"):
                    st.dataframe(
                        filtered_df.sort_values(by=['Date', 'Product', 'Brand']),
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาดระหว่างสร้างไฟล์: {e}")