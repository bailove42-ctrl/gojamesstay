@echo off
title GoJames Booking WebApp
cd /d "%~dp0"
echo ==================================
echo   GoJames Booking WebApp Start
echo ==================================
echo.
echo เว็บนี้เปิดจากมือถือได้ ถ้าอยู่ Wi-Fi เดียวกับคอม
echo ถ้าจะเข้าในมือถือ ให้ดู IPv4 ของคอมด้วยคำสั่ง ipconfig
echo แล้วเปิด http://IPv4:5000
echo.
py app.py

echo.
echo Server stopped.
pause
