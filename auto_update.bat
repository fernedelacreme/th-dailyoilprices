@echo off
cd /d "%~dp0"

:: 1. ดึงโค้ดล่าสุดจาก GitHub ลงมาก่อนเสมอ
git pull --rebase origin main

:: 2. รันสคริปต์ดึงราคาน้ำมัน
python scraper.py

:: 3. เตรียมไฟล์ Parquet
git add data/*.parquet

:: 4. บันทึก Commit
git commit -m "Auto-update prices: %date%"

:: 5. ดันไฟล์ขึ้น GitHub
git push origin main