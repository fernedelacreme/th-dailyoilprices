@echo off
cd /d "%~dp0"

:: 1. รันสคริปต์ดึงราคาน้ำมันเติมไฟล์ Parquet
python scraper.py

:: 2. Stage ไฟล์ Parquet และไฟล์สคริปต์ทั้งหมดก่อน
git add data/*.parquet scraper.py app.py

:: 3. Commit งานในเครื่องไว้ก่อนเสมอ
git commit -m "Auto-update prices: %date%"

:: 4. ดึงข้อมูลล่าสุดจาก GitHub ลงมาผสาน (rebase) หลัง commit แล้ว
git pull --rebase origin main

:: 5. ดันข้อมูลขึ้น GitHub
git push origin main