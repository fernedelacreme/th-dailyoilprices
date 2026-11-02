@echo off
cd /d "%~dp0"

:: 1. รันสคริปต์ดึงราคาน้ำมัน
python scraper.py

:: 2. Stage ไฟล์ทั้งหมดที่ถูกแก้ไข (รวมถึงไฟล์ .bat และ .py)
git add .

:: 3. Commit งานในเครื่องให้เรียบร้อยก่อน
git commit -m "Auto-update prices: %date%"

:: 4. ดึงข้อมูลล่าสุดจาก GitHub มาต่อท้าย
git pull --rebase origin main

:: 5. Push ขึ้น GitHub
git push origin main